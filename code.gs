// ============================================================
// YCT Stock Checker — Google Apps Script (code.gs)
// วิธีติดตั้ง:
// 1. เปิด Google Sheet → Extensions → Apps Script
// 2. วางโค้ดนี้ทับ code.gs
// 3. ตั้ง API Key: Project Settings → Script Properties
//    เพิ่ม key: ANTHROPIC_API_KEY  value: sk-ant-...
// 4. Deploy → New deployment → Web app
//    Execute as: Me | Who has access: Anyone
// ============================================================

const SPREADSHEET_ID = "18VoXGrWL4zsdxCRKICEiJH5Nm-0v2onWXTn45SEYgbY";

const COL_CHECK_DATE = 9;   // col I
const COL_CHECK_STAT = 10;  // col J

// ============================================================
// GET — ดึงข้อมูล stock ทุก sheet
// ============================================================
function doGet(e) {
  try {
    const sheetName = e.parameter.sheet || '';
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = ss.getSheets();
    const targetSheets = sheetName
      ? sheets.filter(s => s.getName() === sheetName)
      : sheets;

    let result = [];
    for (const sheet of targetSheets) {
      const data = sheet.getDataRange().getValues();
      if (data.length < 2) continue;

      let headerRow = 0, colCar = 4, colMachine = 5, colCustomer = 6;
      for (let i = 0; i < Math.min(5, data.length); i++) {
        for (let j = 0; j < data[i].length; j++) {
          const cell = String(data[i][j]).toLowerCase();
          if (cell.includes('หมายเลขรถ') || cell.includes('เลขรถ')) colCar = j;
          if (cell.includes('หมายเลขเครื่อง') || cell.includes('เลขเครื่อง')) { colMachine = j; headerRow = i; }
          if (cell.includes('ชื่อลูกค้า') || cell.includes('ลูกค้า')) colCustomer = j;
        }
      }

      for (let i = headerRow + 1; i < data.length; i++) {
        const machineNo = String(data[i][colMachine] || '').trim();
        if (!machineNo || machineNo.length < 2) continue;
        result.push({
          sheetName: sheet.getName(),
          row: i + 1,
          carNo: String(data[i][colCar] || '').trim(),
          machineNo,
          customer: String(data[i][colCustomer] || '').trim(),
          checkDate: data[i][COL_CHECK_DATE - 1] ? String(data[i][COL_CHECK_DATE - 1]) : '',
          checkStatus: data[i][COL_CHECK_STAT - 1] ? String(data[i][COL_CHECK_STAT - 1]) : '',
        });
      }
    }

    return jsonResponse({ ok: true, data: result, total: result.length });
  } catch(err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ============================================================
// POST — รับทั้ง OCR (รูปภาพ) และ writeback (ผลลัพธ์)
// action=ocr  → base64 image → Anthropic API → serial
// action=write → บันทึกผลลง Sheet
// ============================================================
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action || 'write';

    if (action === 'ocr') {
      return handleOCR(payload);
    } else {
      return handleWrite(payload);
    }
  } catch(err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ============================================================
// OCR — ส่งรูปไป Anthropic แล้วคืน serial
// ============================================================
function handleOCR(payload) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return jsonResponse({ ok: false, error: 'ไม่พบ ANTHROPIC_API_KEY ใน Script Properties' });

  const { imageBase64, mediaType } = payload;
  if (!imageBase64) return jsonResponse({ ok: false, error: 'ไม่มีข้อมูลรูปภาพ' });

  const requestBody = {
    model: 'claude-sonnet-4-20250514',
    max_tokens: 300,
    messages: [{
      role: 'user',
      content: [
        {
          type: 'image',
          source: { type: 'base64', media_type: mediaType || 'image/jpeg', data: imageBase64 }
        },
        {
          type: 'text',
          text: `Read the Yanmar equipment nameplate in this photo.
Extract exactly:
- MODEL: model name (e.g. YM358R, EF393T)
- SERIAL: full MODEL SERIAL NO. string (e.g. YMJS0057CPLK50488)
- SUFFIX: trailing numeric digits from serial (e.g. 50488)

Return ONLY JSON, no markdown:
{"model":"...","serial":"...","suffix":"..."}
Use null if unreadable.`
        }
      ]
    }]
  };

  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  });

  const respCode = response.getResponseCode();
  const respText = response.getContentText();

  if (respCode !== 200) {
    return jsonResponse({ ok: false, error: `Anthropic error ${respCode}: ${respText.substring(0, 200)}` });
  }

  const respData = JSON.parse(respText);
  const text = respData.content?.find(c => c.type === 'text')?.text || '';
  const clean = text.replace(/```json|```/g, '').trim();

  try {
    const parsed = JSON.parse(clean);
    return jsonResponse({ ok: true, result: parsed });
  } catch(e) {
    return jsonResponse({ ok: false, error: 'Parse error: ' + clean.substring(0, 100) });
  }
}

// ============================================================
// WRITE — บันทึกผลลง Sheet + ทำสี
// ============================================================
function handleWrite(payload) {
  const results = payload.results || [];
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let written = 0, errors = [];

  // จัดกลุ่มตาม sheet สำหรับทำสี
  const bySheet = {};

  for (const r of results) {
    try {
      const sheet = ss.getSheetByName(r.sheetName);
      if (!sheet) { errors.push(`Sheet "${r.sheetName}" ไม่พบ`); continue; }
      const rowNum = parseInt(r.row);
      if (!rowNum) { errors.push(`row ไม่ถูกต้อง: ${r.row}`); continue; }

      sheet.getRange(rowNum, COL_CHECK_DATE).setValue(
        r.checkDate || new Date().toLocaleDateString('th-TH')
      );
      sheet.getRange(rowNum, COL_CHECK_STAT).setValue({
        'instock':  '✅ ตรวจพบ - มีสต๊อก',
        'sold':     '🔴 ตรวจพบ - ขายแล้ว',
        'notfound': '⚠️ ไม่พบรูปคู่',
        'error':    '❌ อ่าน Serial ไม่ได้',
      }[r.status] || r.status);

      if (!bySheet[r.sheetName]) bySheet[r.sheetName] = [];
      bySheet[r.sheetName].push(r);
      written++;
    } catch(rowErr) {
      errors.push(`แถว ${r.row}: ${rowErr.message}`);
    }
  }

  // ทำสีแถว
  const colorMap = { instock:'#d9ead3', sold:'#fce8e6', notfound:'#fff2cc', error:'#efefef' };
  for (const [sheetName, rows] of Object.entries(bySheet)) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;
    const lastCol = sheet.getLastColumn();
    for (const r of rows) {
      const rowNum = parseInt(r.row);
      if (!rowNum) continue;
      sheet.getRange(rowNum, 1, 1, lastCol).setBackground(colorMap[r.status] || '#ffffff');
    }
  }

  return jsonResponse({ ok: true, written, errors });
}

// ============================================================
// Helper
// ============================================================
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

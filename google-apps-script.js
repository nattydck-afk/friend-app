/**
 * =======================================================================
 * FRIEND APP (แก๊งค์ลูกหมู 🐷✈️) — GOOGLE APPS SCRIPT WEB APP BACKEND
 * =======================================================================
 * 
 * วิธีติดตั้งเพื่อเชื่อมต่อการเขียนข้อมูลจริง 2-way เข้า Google Sheet:
 * 1. เปิด Google Sheet: https://docs.google.com/spreadsheets/d/1wc5mk9XH7EUVBJeN30J2wlt7Wqk4setS1XFi-yKxyHs/edit
 * 2. ไปที่เมนู "Extensions" (ส่วนขยาย) > "Apps Script"
 * 3. นำโค้ดนี้ทั้งหมดไปวางแทนที่โค้ดเดิมใน Editor
 * 4. กดปุ่ม "Deploy" (การทำให้ใช้งานได้) > "New deployment" (การทำให้ใช้งานได้ใหม่)
 * 5. เลือกประเภท "Web app" (เว็บแอป)
 * 6. ตั้งค่า:
 *    - Description: FriendApp Backend
 *    - Execute as: "Me" (ฉัน)
 *    - Who has access: "Anyone" (ทุกคน)  <-- สำคัญมาก! เพื่อให้เว็บแอปบันทึกข้อมูลได้
 * 7. กด "Deploy" แล้ว Copy "Web App URL" (ที่ลงท้ายด้วย /exec)
 * 8. นำ URL มาใส่ใน Friend App หน้าเว็บ (มีปุ่ม ⚙️ ตั้งค่าเชื่อมต่อ Google Sheet)
 */

const SPREADSHEET_ID = '1wc5mk9XH7EUVBJeN30J2wlt7Wqk4setS1XFi-yKxyHs';

function getSpreadsheet() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (err) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
}

function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const action = params.action || 'getAll';

  try {
    const ss = getSpreadsheet();

    if (action === 'getAll') {
      const sumSheet = ss.getSheetByName('SumView');
      const manageSheet = ss.getSheetByName('จัดการรายการ');
      const collectSheet = ss.getSheetByName('เรียกเก็บ');

      return respondJSON({
        success: true,
        sumView: sumSheet ? getSheetData(sumSheet) : [],
        manageItems: manageSheet ? getSheetData(manageSheet) : [],
        collectList: collectSheet ? getSheetData(collectSheet) : []
      });
    }

    const sheetName = params.sheet || 'เรียกเก็บ';
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return respondJSON({ success: false, error: 'Sheet not found: ' + sheetName });
    }

    return respondJSON({
      success: true,
      sheet: sheetName,
      data: getSheetData(sheet)
    });
  } catch (error) {
    return respondJSON({ success: false, error: error.toString() });
  }
}

function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    const action = payload.action;
    const sheetName = payload.sheet || 'เรียกเก็บ';
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return respondJSON({ success: false, error: 'Sheet not found: ' + sheetName });
    }

    // 1. APPEND: เพิ่มรายการใหม่ (เช่น เพิ่มรายการใน "เรียกเก็บ" หรือ "จัดการรายการ")
    if (action === 'append') {
      const rows = Array.isArray(payload.rows) ? payload.rows : [payload.data];
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

      rows.forEach(r => {
        if (Array.isArray(r)) {
          sheet.appendRow(r);
        } else {
          const newRow = headers.map(h => r[h] !== undefined ? r[h] : '');
          sheet.appendRow(newRow);
        }
      });

      return respondJSON({ success: true, message: 'Rows appended successfully', count: rows.length });
    }

    // 2. UPDATE: อัปเดตข้อมูล (เช่น แนบสลิป, เปลี่ยนสถานะเป็นโอนแล้ว, แก้ไขยอด)
    if (action === 'update') {
      const matchTrip = payload.matchTrip;
      const matchMember = payload.matchMember;
      const updateData = payload.data || {};

      const values = sheet.getDataRange().getValues();
      if (values.length <= 1) {
        return respondJSON({ success: false, error: 'No data rows found' });
      }

      const headers = values[0];
      const tripColIdx = headers.indexOf('ชื่อรายการเรียกเก็บ') !== -1 ? headers.indexOf('ชื่อรายการเรียกเก็บ') : headers.indexOf('ทริป');
      const nameColIdx = headers.indexOf('ชื่อ');

      let updatedCount = 0;

      for (let r = 1; r < values.length; r++) {
        const rowTrip = String(values[r][tripColIdx]).trim();
        const rowName = String(values[r][nameColIdx]).trim();

        if (rowTrip === String(matchTrip).trim() && rowName === String(matchMember).trim()) {
          // Update columns
          Object.keys(updateData).forEach(colName => {
            const cIdx = headers.indexOf(colName);
            if (cIdx !== -1) {
              sheet.getRange(r + 1, cIdx + 1).setValue(updateData[colName]);
            }
          });
          updatedCount++;
        }
      }

      return respondJSON({
        success: updatedCount > 0,
        message: updatedCount > 0 ? 'Updated ' + updatedCount + ' row(s)' : 'No matching row found'
      });
    }

    return respondJSON({ success: false, error: 'Unknown action: ' + action });

  } catch (error) {
    return respondJSON({ success: false, error: error.toString() });
  }
}

function getSheetData(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0].map(h => String(h).trim());
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const rowObj = { _rowIndex: i + 1 };
    let hasValue = false;
    for (let j = 0; j < headers.length; j++) {
      rowObj[headers[j]] = data[i][j];
      if (data[i][j] !== '' && data[i][j] !== null && data[i][j] !== undefined) {
        hasValue = true;
      }
    }
    if (hasValue) rows.push(rowObj);
  }
  return rows;
}

function respondJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * =========================================================================
 * Manage Money - Real-Time Google Sheets Sync Script
 * สคริปต์เชื่อมต่อข้อมูลแบบเรียลไทม์ระหว่าง Web App มือถือ กับ Google Sheets
 * =========================================================================
 * 
 * วิธีติดตั้ง (ทำเพียงครั้งเดียว ใช้เวลา 1-2 นาที):
 * 1. เปิดไฟล์ Google Sheets ของคุณ
 * 2. ไปที่เมนูด้านบน คลิก "ส่วนขยาย" (Extensions) > เลือก "Apps Script"
 * 3. ลบโค้ดเดิมทั้งหมดในหน้าต่าง แล้วคัดลอกโค้ดนี้ทั้งหมดไปวางแทนที่
 * 4. คลิกปุ่ม "บันทึก" (รูปแผ่นดิสก์ 💾 หรือกด Ctrl+S)
 * 5. คลิกปุ่มสีน้ำเงินด้านขวาบน "ทำให้ใช้งานได้" (Deploy) > เลือก "การทำให้ใช้งานได้ใหม่" (New deployment)
 * 6. ตรงรูปฟันเฟือง ⚙️ ให้เลือกประเภทเป็น "เว็บแอป" (Web app)
 *    - คำอธิบาย: Manage Money Real-time Sync
 *    - ดำเนินการในฐานะ (Execute as): "ฉัน" (Me - บัญชีของคุณ)
 *    - ผู้ที่มีสิทธิ์เข้าถึง (Who has access): "ทุกคน" (Anyone) **(สำคัญมาก เพื่อให้มือถือส่งข้อมูลได้)**
 * 7. คลิกปุ่ม "ทำให้ใช้งานได้" (Deploy) > หากมีหน้าต่างขอสิทธิ์ ให้กด "ตรวจสอบสิทธิ์" (Authorize) แล้วกดยืนยัน
 * 8. คุณจะได้ "URL ของเว็บแอป" (Web App URL) ที่ลงท้ายด้วย /exec
 * 9. นำ URL นั้นมาวางในช่อง "ตั้งค่า Google Sheets" ในแอป Manage money บนมือถือของคุณ
 * =========================================================================
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  // รอคิวล็อกสูงสุด 15 วินาที เพื่อความปลอดภัยในกรณีมีหลายรายการเข้ามาพร้อมกัน
  lock.tryLock(15000);
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // ค้นหาชีต บันทึกรายจ่าย และ ภาพรวม
    var sheetExpenses = ss.getSheetByName("บันทึกรายจ่าย");
    var sheetOverview = ss.getSheetByName("ภาพรวม & งบประมาณ");
    
    // Fallback หากผู้ใช้เปลี่ยนชื่อชีต
    if (!sheetExpenses) {
      sheetExpenses = ss.getSheets()[1] || ss.getSheets()[0];
    }
    if (!sheetOverview) {
      sheetOverview = ss.getSheets()[0];
    }
    
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({ status: "error", message: "No data received" });
    }
    
    var data = JSON.parse(e.postData.contents);
    var action = data.action || "addExpense";
    
    // -----------------------------------------------------------------
    // 1. เพิ่มรายการรายจ่ายใหม่ลงตารางบันทึกรายจ่าย
    // -----------------------------------------------------------------
    if (action === "addExpense") {
      var dateStr = data.date || Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
      var title = data.title || "รายการใช้จ่าย";
      var pocketName = data.pocketName || getPocketLabel(data.category);
      var subCategory = data.subCategory || "ทั่วไป";
      var amount = Number(data.amount) || 0;
      var paymentMethod = data.paymentMethod || "สแกน QR / โอนเงิน";
      var note = data.note || "บันทึกจากมือถือ";
      
      sheetExpenses.appendRow([
        dateStr,
        title,
        pocketName,
        subCategory,
        amount,
        paymentMethod,
        note
      ]);
      
      // จัดรูปแบบตัวเลขคอลัมน์เงิน
      var lastRow = sheetExpenses.getLastRow();
      sheetExpenses.getRange(lastRow, 5).setNumberFormat("#,##0.00");
      
      return createJsonResponse({
        status: "success",
        message: "บันทึกรายจ่ายลง Google Sheets สำเร็จ",
        rowAdded: lastRow
      });
    }
    
    // -----------------------------------------------------------------
    // 2. อัปเดตรายรับสุทธิประจำเดือนในชีตภาพรวม
    // -----------------------------------------------------------------
    if (action === "setIncome") {
      var incomeVal = Number(data.income) || 35000;
      sheetOverview.getRange("C4").setValue(incomeVal);
      return createJsonResponse({
        status: "success",
        message: "อัปเดตรายได้สุทธิใน Google Sheets สำเร็จ",
        income: incomeVal
      });
    }
    
    // -----------------------------------------------------------------
    // 3. ซิงค์ข้อมูลประวัติทั้งหมดขึ้นชีต (Sync All)
    // -----------------------------------------------------------------
    if (action === "syncAll") {
      var items = data.expenses || [];
      if (data.mode === "replace") {
        var lastRow = sheetExpenses.getLastRow();
        if (lastRow >= 2) {
          sheetExpenses.getRange(2, 1, lastRow - 1, 7).clearContent();
        }
      }
      
      for (var i = 0; i < items.length; i++) {
        var it = items[i];
        sheetExpenses.appendRow([
          it.date,
          it.title,
          it.pocketName || getPocketLabel(it.category),
          it.subCategory || "ทั่วไป",
          Number(it.amount) || 0,
          it.paymentMethod || "สแกน QR / โอนเงิน",
          it.note || "บันทึกจากมือถือ"
        ]);
      }
      
      return createJsonResponse({
        status: "success",
        message: "ซิงค์ประวัติทั้งหมดสำเร็จ",
        count: items.length
      });
    }
    
    return createJsonResponse({ status: "ignored", message: "Action not supported" });
    
  } catch (err) {
    return createJsonResponse({ status: "error", message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetExpenses = ss.getSheetByName("บันทึกรายจ่าย") || ss.getSheets()[1] || ss.getSheets()[0];
    var sheetOverview = ss.getSheetByName("ภาพรวม & งบประมาณ") || ss.getSheets()[0];
    
    var income = sheetOverview.getRange("C4").getValue();
    var lastRow = sheetExpenses.getLastRow();
    var expenses = [];
    
    if (lastRow >= 2) {
      var values = sheetExpenses.getRange(2, 1, lastRow - 1, 7).getValues();
      for (var i = 0; i < values.length; i++) {
        var row = values[i];
        if (row[0] || row[1] || row[4]) {
          expenses.push({
            id: i + 1,
            date: formatDateValue(row[0]),
            title: String(row[1] || ""),
            pocketName: String(row[2] || ""),
            category: parseCategoryKey(row[2]),
            subCategory: String(row[3] || ""),
            amount: Number(row[4]) || 0,
            paymentMethod: String(row[5] || ""),
            note: String(row[6] || "")
          });
        }
      }
    }
    
    return createJsonResponse({
      status: "success",
      title: "Manage Money Google Sheets API",
      income: Number(income) || 35000,
      totalExpensesCount: expenses.length,
      expenses: expenses
    });
  } catch (err) {
    return createJsonResponse({ status: "error", message: err.toString() });
  }
}

function getPocketLabel(cat) {
  if (cat === "needs") return "1. จำเป็น (Needs)";
  if (cat === "wealth") return "2. ออม & ลงทุน (Wealth)";
  if (cat === "wants") return "3. ใช้ชีวิต (Wants)";
  if (cat === "growth") return "4. พัฒนาตน (Growth)";
  return "1. จำเป็น (Needs)";
}

function parseCategoryKey(str) {
  var s = String(str || "");
  if (s.indexOf("จำเป็น") !== -1 || s.indexOf("Needs") !== -1) return "needs";
  if (s.indexOf("ออม") !== -1 || s.indexOf("Wealth") !== -1) return "wealth";
  if (s.indexOf("ใช้ชีวิต") !== -1 || s.indexOf("Wants") !== -1) return "wants";
  if (s.indexOf("พัฒนาตน") !== -1 || s.indexOf("Growth") !== -1) return "growth";
  return "wants";
}

function formatDateValue(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, "GMT+7", "yyyy-MM-dd");
  }
  return String(val || "");
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

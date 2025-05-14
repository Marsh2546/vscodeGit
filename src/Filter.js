// /**
//  * อัปเดตรายการเดือนและปีจากข้อมูลในชีต 'RawData'
//  * และสร้าง dropdown ให้เลือกเดือนและปีแบบ dynamic ในชีต 'Dashboard'
//  */
// function updateMonthYearFilter() {
//   // เข้าถึง Spreadsheet และ Sheet ที่ต้องการใช้งาน
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const rawSheet = ss.getSheetByName("คำสั่งซื้อทั้งหมด");     // ชีตข้อมูลดิบ
//   const dashboard = ss.getSheetByName("Dashboard");   // ชีตที่แสดง dashboard

//   // ดึงค่าวันที่จากคอลัมน์ A ตั้งแต่แถว 2 ลงไปในชีต RawData
//   const dates = rawSheet.getRange("AC2:AC" + rawSheet.getLastRow()).getValues().flat();

//   const months = []; // เก็บเดือนแบบไม่ซ้ำ เช่น "01", "02"
//   const years = [];  // เก็บปีแบบไม่ซ้ำ เช่น "2023", "2024"

//   // วนลูปผ่านรายการวันที่ และแยกเดือน/ปีออกมาแบบไม่ซ้ำ
//   dates.forEach(date => {
//     if (date instanceof Date) { // ตรวจสอบว่าเป็นประเภท Date
//       const month = ("0" + (date.getMonth() + 1)).slice(-2); // คืนค่าเดือนแบบ 2 หลัก
//       const year = date.getFullYear().toString();
//       if (!months.includes(month)) months.push(month);
//       if (!years.includes(year)) years.push(year);
//     }
//   });

//   // เรียงเดือนและปีจากน้อยไปมาก
//   months.sort();
//   years.sort();

//   // ล้างข้อมูลเก่าของรายการเดือนและปีในคอลัมน์ G และ H ของ Dashboard
//   dashboard.getRange("U2:U").clearContent();
//   dashboard.getRange("V2:V").clearContent();

//   // ถ้ามีเดือน ให้เขียนลงใน G2:G
//   if (months.length > 0) {
//     dashboard.getRange("U2:U" + (months.length + 1)).setValues(months.map(m => [m]));
//   }

//   // ถ้ามีปี ให้เขียนลงใน H2:H
//   if (years.length > 0) {
//     dashboard.getRange("V2:V" + (years.length + 1)).setValues(years.map(y => [y]));
//   }

//   // สร้าง dropdown ให้เซลล์ B1 (ปี) และ C1 (เดือน)
//   const monthCell = dashboard.getRange("D3");
//   const yearCell = dashboard.getRange("D2");

//   // ถ้ามีเดือน ให้สร้าง dropdown อ้างอิงช่วง G2:G...
//   if (months.length > 0) {
//     monthCell.setDataValidation(
//       SpreadsheetApp.newDataValidation()
//         .requireValueInRange(dashboard.getRange("U2:U" + (months.length + 1)), true)
//         .build()
//     );
//   } else {
//     monthCell.clearDataValidations(); // ถ้าไม่มีเดือน ให้ลบ dropdown เดิมทิ้ง
//   }

//   // ถ้ามีปี ให้สร้าง dropdown อ้างอิงช่วง H2:H...
//   if (years.length > 0) {
//     yearCell.setDataValidation(
//       SpreadsheetApp.newDataValidation()
//         .requireValueInRange(dashboard.getRange("V2:V" + (years.length + 1)), true)
//         .build()
//     );
//   } else {
//     yearCell.clearDataValidations(); // ถ้าไม่มีปี ให้ลบ dropdown เดิมทิ้ง
//   }
// }

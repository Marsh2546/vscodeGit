/**
 * ฟังก์ชัน updateDashboardData
 * ใช้เพื่ออัปเดตข้อมูลยอดขาย ต้นทุน และกำไรใน Dashboard ตามเดือน/ปี ที่ผู้ใช้เลือก
 */
function updateDashboardData() {
  // เข้าถึง spreadsheet และกำหนดชีตข้อมูลดิบและชีต dashboard
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("คำสั่งซื้อทั้งหมด");
  const dashboard = ss.getSheetByName("Dashboard");

  // อ่านค่าปีและเดือนที่ผู้ใช้เลือกจากเซลล์ B1 และ C1 บนชีต Dashboard
  const selectedYear = dashboard.getRange("B1").getValue().toString();
  const selectedMonth = dashboard.getRange("C1").getValue().toString();

  // ดึงข้อมูลช่วง A2:C จากชีต RawData (A = วันที่, B = ยอดขาย, C = ต้นทุน)
  const data = rawSheet.getRange("A2:C" + rawSheet.getLastRow()).getValues();

  let totalSales = 0; // ตัวแปรเก็บยอดขายรวม
  let totalCost = 0;  // ตัวแปรเก็บต้นทุนรวม

  // วนลูปข้อมูลทีละแถว เพื่อรวมยอดขายและต้นทุนตามเดือน/ปีที่เลือก
  data.forEach(row => {
    const date = row[0]; // วันที่ในคอลัมน์ A
    const sales = Number(row[1]); // ยอดขายในคอลัมน์ B
    const cost = Number(row[2]);  // ต้นทุนในคอลัมน์ C

    if (date instanceof Date) {
      const month = ("0" + (date.getMonth() + 1)).slice(-2); // แปลงเดือนเป็นเลข 2 หลัก เช่น "01"
      const year = date.getFullYear().toString();

      // เช็คว่าเดือนและปีตรงกับที่ผู้ใช้เลือกหรือไม่
      if (month === selectedMonth && year === selectedYear) {
        totalSales += sales;
        totalCost += cost;
      }
    }
  });

  const profit = totalSales - totalCost; // คำนวณกำไร

  // เขียนค่าไปยัง Dashboard:
  // D2 = ยอดขายรวม, E2 = ต้นทุนรวม, F2 = กำไรรวม
  dashboard.getRange("D2").setValue(totalSales);
  dashboard.getRange("E2").setValue(totalCost);
  dashboard.getRange("F2").setValue(profit);
}

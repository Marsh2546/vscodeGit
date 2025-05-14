function generateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("คำสั่งซื้อทั้งหมด");
  const dashSheet = ss.getSheetByName("Dashboard Info");
  if (!dashSheet) {
    SpreadsheetApp.getUi().alert("ไม่พบชีทชื่อ 'Dashboard'");
    return;
  }

  const data = rawSheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const dateCol = headers.indexOf("วันที่สร้างคำสั่งซื้อ");
  const totalCol = headers.indexOf("ยอดรวมสุทธิคำสั่งซื้อ");
  const typeCol = headers.indexOf("ประเภทสินค้า");
  const styleCol = headers.indexOf("ลักษณะ");
  const channelCol = headers.indexOf("ช่องทาง");
  const returnCol = headers.indexOf("Cancelation/Return Type"); // ✅ เพิ่มตรงนี้

  const monthlyStats = {};
  const productSales = {};
  const channelSales = {};

  // ✅ ใช้แค่รอบเดียว ไม่ซ้อนกัน
  rows.forEach(row => {
    const rawDate = new Date(row[dateCol]);
    if (isNaN(rawDate)) return;

    // 🚫 ข้ามรายการคืนสินค้า
    if (row[returnCol] && row[returnCol].toString().trim() !== "") return;

    const yearMonth = `${rawDate.getFullYear()}-${('0' + (rawDate.getMonth() + 1)).slice(-2)}`;
    const total = parseFloat(row[totalCol]) || 0;

    const productType = row[typeCol] || "";
    const productStyle = row[styleCol] || "";
    const productKey = `${productType} / ${productStyle}`;
    const channel = row[channelCol];

    // Monthly Summary
    if (!monthlyStats[yearMonth]) {
      monthlyStats[yearMonth] = { sales: 0, count: 0 };
    }
    monthlyStats[yearMonth].sales += total;
    monthlyStats[yearMonth].count += 1;

    // Product Summary
    if (!productSales[productKey]) productSales[productKey] = 0;
    productSales[productKey] += total;

    // Channel Summary
    if (!channelSales[channel]) channelSales[channel] = 0;
    channelSales[channel] += total;
  });

  // ล้างข้อมูลเก่า
  dashSheet.clearContents();

  let row = 1;
  dashSheet.getRange(row++, 1).setValue("📈 สรุปข้อมูลคำสั่งซื้อ (เพิ่มเมื่อ: " + new Date().toLocaleString() + ")");

  dashSheet.getRange(row++, 1).setValue("ยอดขายรวมรายเดือน:");
  dashSheet.getRange(row++, 1, 1, 3).setValues([["เดือน", "ยอดขายรวม", "จำนวนคำสั่งซื้อ"]]);

  Object.keys(monthlyStats).sort().forEach(key => {
    const stats = monthlyStats[key];
    dashSheet.getRange(row++, 1, 1, 3).setValues([[key, stats.sales, stats.count]]);
  });

  row += 2;
  dashSheet.getRange(row++, 1).setValue("สินค้าที่ขายดีสุด 5 อันดับ:");
  dashSheet.getRange(row++, 1, 1, 2).setValues([["สินค้า", "ยอดขาย"]]);

  const topProducts = Object.entries(productSales)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);

  topProducts.forEach(([product, total]) => {
    dashSheet.getRange(row++, 1, 1, 2).setValues([[product, total]]);
  });

  row += 2;
  dashSheet.getRange(row++, 1).setValue("ยอดขายแยกตามช่องทาง:");
  dashSheet.getRange(row++, 1, 1, 2).setValues([["ช่องทาง", "ยอดขาย"]]);

  Object.entries(channelSales).forEach(([channel, total]) => {
    dashSheet.getRange(row++, 1, 1, 2).setValues([[channel, total]]);
  });
}

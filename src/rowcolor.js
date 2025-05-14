function applyConditionalFormatting() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("สั่งของ");
  const range = sheet.getRange("A2:P"); // ปรับช่วงข้อมูลตามจริง
  const rules = [];

  // ลบกฎเดิมทั้งหมดก่อน
  sheet.clearConditionalFormatRules();

  // รายชื่อสินค้าและสี
  const items = [
    { text: "แผ่นรองเมาส์", color: "#CFEAFE" },      // เขียวอ่อน
    { text: "แผ่นรองเมาส์ FPS", color: "#BFD7DD" },      // เขียวอ่อน
    { text: "สาย keyboard", color: "#FFD6B3" },       // ส้มอ่อน
    { text: "สาย keyboard High Speed", color: "#FFE699" },  // แดงอ่อน
    { text: "Lightbar Monitor", color: "#D9FAD3" },  // ฟ้าอ่อน
    { text: "Microphone Arm Lowprofile", color: "#FFC6C6" },    // แดงอ่อน
    { text: "แผ่นรองเมาส์ FPS (TYPE 99)", color: "#E8D9F9" }    // แดงอ่อน
  ];

  // สร้างกฎแต่ละอัน
  for (let i = 0; i < items.length; i++) {
    const formula = `=$F2="${items[i].text}"`;
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground(items[i].color)
      .setRanges([range])
      .build();
    rules.push(rule);
  }

  // ใช้กฎทั้งหมดกับชีท
  sheet.setConditionalFormatRules(rules);
}
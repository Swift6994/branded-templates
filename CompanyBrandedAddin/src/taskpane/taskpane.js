Office.onReady(info => {
  if (info.host) {
    document.getElementById("insert").onclick = insertCompanyTemplate;
  }
});

async function insertCompanyTemplate() {
  const host = Office.context.host;
  const status = document.getElementById("status");
  status.textContent = "Inserting template...";

  try {
    if (host === Office.HostType.Word) {
      await Word.run(async context => {
        const body = context.document.body;
        body.clear();
        const title = body.insertParagraph("🚀 ACME Corp — Internal Report", "Start");
        title.font.set({
          name: "Calibri Light",
          size: 28,
          color: "#004B8D",
          bold: true
        });
        const date = body.insertParagraph("Generated: " + new Date().toLocaleDateString(), "After");
        date.font.set({
          name: "Calibri",
          size: 12,
          color: "gray"
        });
        const separator = body.insertParagraph("─────────────────────────────", "After");
        separator.font.color = "#004B8D";
        body.insertParagraph("Executive Summary:\n\n", "After");
        await context.sync();
      });

    } else if (host === Office.HostType.Excel) {
      await Excel.run(async context => {
        const sheet = context.workbook.getActiveWorksheet();
        const header = sheet.getRange("A1:C1");
        header.values = [["ACME Corp Report", "Prepared By", "Date"]];
        header.format.fill.color = "#004B8D";
        header.format.font.color = "white";
        header.format.font.bold = true;

        const data = sheet.getRange("A2:C2");
        data.values = [["Quarterly Summary", "Finance Dept", new Date().toLocaleDateString()]];
        data.format.autofitColumns();
        sheet.freezePanes.freezeRows(1);
        await context.sync();
      });

    } else if (host === Office.HostType.PowerPoint) {
      await PowerPoint.run(async context => {
        const slide = context.presentation.slides.getItemAt(0);
        slide.shapes.addTextBox("🚀 ACME Corp Presentation")
          .textFrame.textRange.font.set({
            bold: true,
            size: 36,
            color: "#004B8D"
          });
        slide.shapes.addTextBox("Innovate. Empower. Deliver.")
          .textFrame.textRange.font.set({
            italic: true,
            size: 20,
            color: "gray"
          });
        await context.sync();
      });

    } else {
      throw new Error("Unsupported host: " + host);
    }

    status.textContent = "✅ Branded template inserted successfully!";
  } catch (error) {
    console.error(error);
    status.textContent = "❌ Error: " + error.message;
  }
}

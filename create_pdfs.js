const XLSX = require("xlsx");
const PDFKit = require("pdfkit");
const fs = require("fs");
const path = require("path");

// Create the 'pdfs' folder if it doesn't exist
const pdfsDir = path.join(__dirname, "pdfs");
if (!fs.existsSync(pdfsDir)) {
  fs.mkdirSync(pdfsDir);
}

// Read the Excel file
const workbook = XLSX.readFile("POs.xlsx"); // Replace with your Excel file name
const sheetName = workbook.SheetNames[0]; // Get the first sheet
const sheet = workbook.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(sheet); // Convert sheet to JSON

// Group data by PO number
const groupedData = {};
data.forEach((row) => {
  const po = row["PO"];
  if (!groupedData[po]) {
    groupedData[po] = [];
  }
  groupedData[po].push(row);
});

// Function to draw a table row
function drawTableRow(doc, y, rowData, colWidths, startX) {
  let x = startX;
  rowData.forEach((cell, i) => {
    doc.text(cell.toString(), x, y, { width: colWidths[i], align: "left" });
    x += colWidths[i];
  });
}

// Create a PDF for each PO
Object.keys(groupedData).forEach((po) => {
  const rows = groupedData[po];
  const pdfPath = path.join(pdfsDir, `PO_${po}.pdf`);

  // Create a new PDF document
  const doc = new PDFKit({ margin: 30 });
  doc.pipe(fs.createWriteStream(pdfPath));

  // Add a title
  doc.fontSize(14).text(`PO Number: ${po}`, { align: "center" });
  doc.moveDown(1);

  // Define table properties
  const headers = Object.keys(rows[0]);
  const colWidths = headers.map(() => 80); // Fixed width for each column (adjust as needed)
  const startX = 30;
  let y = doc.y;

  // Draw table header
  doc.fontSize(10).font("Helvetica-Bold");
  drawTableRow(doc, y, headers, colWidths, startX);
  y += 20;

  // Draw header underline
  doc
    .moveTo(startX, y - 5)
    .lineTo(startX + colWidths.reduce((a, b) => a + b, 0), y - 5)
    .stroke();
  y += 5;

  // Draw table rows
  doc.font("Helvetica");
  rows.forEach((row) => {
    const rowData = headers.map((header) => row[header] || "");
    drawTableRow(doc, y, rowData, colWidths, startX);
    y += 20;

    // Start a new page if needed
    if (y > 700) {
      doc.addPage();
      y = 50;
      // Redraw header on new page
      doc.font("Helvetica-Bold");
      drawTableRow(doc, y, headers, colWidths, startX);
      y += 20;
      doc
        .moveTo(startX, y - 5)
        .lineTo(startX + colWidths.reduce((a, b) => a + b, 0), y - 5)
        .stroke();
      y += 5;
      doc.font("Helvetica");
    }
  });

  // Finalize the PDF
  doc.end();
  console.log(`Created PDF: ${pdfPath}`);
});

console.log('All PDFs have been created in the "pdfs" folder!');

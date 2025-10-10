document.getElementById("processBtn").addEventListener("click", async () => {
  console.log("Button clicked!");
  const pdfFile = document.getElementById("pdfInput").files[0];
  const excelFile = document.getElementById("excelInput").files[0];
  const outputDiv = document.getElementById("output");
  outputDiv.innerHTML = "";

  if (!pdfFile || !excelFile) {
    alert("Please select both a PDF and an Excel file.");
    return;
  }

  try {
    // --- Step 1: Read Excel ---
    const excelBuffer = await excelFile.arrayBuffer();
    const workbook = XLSX.read(excelBuffer, { type: "array" });
    const firstSheet = workbook.SheetNames[0];
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet]);

    // --- Step 2: Process like pandas ---
    // Sort by DATE, From
    data.sort((a, b) => {
      const dateA = new Date(a["DATE"]);
      const dateB = new Date(b["DATE"]);
      if (dateA - dateB !== 0) return dateA - dateB;
      return (a["From"] || "").localeCompare(b["From"] || "");
    });

    // Grouping logic
    const groupKeys = ["SCHOOL", "TUTOR", "Session ID", "HRS", "DATE", "Duration", "Hourly Rate", "TOTAL $ P"];
    const grouped = {};
    for (const row of data) {
      const key = groupKeys.map(k => row[k]).join("|");
      if (!grouped[key]) grouped[key] = { ...row, NAME: new Set() };
      grouped[key].NAME.add(row["NAME"]);
    }

    let groupedArr = Object.values(grouped).map(obj => ({
      ...obj,
      NAME: Array.from(obj.NAME).join("\n")
    }));

    // Reorder columns: move NAME to position 2
    const cols = Object.keys(groupedArr[0]);
    const lastCol = cols.pop();
    cols.splice(2, 0, lastCol);

    // --- Step 3: Render table to canvas ---
    const tableCanvas = document.createElement("canvas");
    const ctx = tableCanvas.getContext("2d");
    ctx.font = "12px Arial";
    const rowHeight = 18;
    const colWidths = cols.map(() => 100);
    const tableWidth = colWidths.reduce((a, b) => a + b, 0);
    const tableHeight = (groupedArr.length + 1) * rowHeight + 40;

    tableCanvas.width = tableWidth;
    tableCanvas.height = tableHeight;

    ctx.fillStyle = "white";
    ctx.fillRect(0, 0, tableWidth, tableHeight);
    ctx.fillStyle = "black";

    // Header
    cols.forEach((c, i) => {
      ctx.fillText(c, 10 + i * 100, 20);
    });

    // Rows
    groupedArr.forEach((r, rowIndex) => {
      cols.forEach((c, colIndex) => {
        const text = String(r[c] ?? "");
        const lines = text.split("\n");
        lines.forEach((line, j) => {
          ctx.fillText(line, 10 + colIndex * 100, 40 + rowIndex * rowHeight + j * 14);
        });
      });
    });

    // Convert canvas to image (for PDF page)
    const tableImgUrl = tableCanvas.toDataURL("image/png");

    // --- Step 4: Load PDF, keep first page ---
    const pdfBytes = await pdfFile.arrayBuffer();
    const originalPdf = await PDFLib.PDFDocument.load(pdfBytes);
    const newPdf = await PDFLib.PDFDocument.create();

    const [firstPage] = await newPdf.copyPages(originalPdf, [0]);
    newPdf.addPage(firstPage);

    // --- Step 5: Add Excel table image as pages ---
    const img = await newPdf.embedPng(tableImgUrl);
    const page = newPdf.addPage([img.width, img.height]);
    page.drawImage(img, { x: 0, y: 0, width: img.width, height: img.height });

    // --- Step 6: Save and download ---
    const mergedBytes = await newPdf.save();
    const blob = new Blob([mergedBytes], { type: "application/pdf" });
    const url = URL.createObjectURL(blob);

    const link = document.createElement("a");
    link.href = url;
    link.download = "merged.pdf";
    link.textContent = "Download Combined PDF";
    link.classList.add("download-link");
    outputDiv.appendChild(link);
  } catch (err) {
    console.error(err);
    alert("Error processing files.");
  }
});

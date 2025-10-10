document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("processBtn").addEventListener("click", async () => {
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

      if (!data.length) {
        alert("Excel file is empty or not formatted properly.");
        return;
      }

      // --- Step 2: Process like pandas ---
      data.sort((a, b) => {
        const dateA = new Date(a["DATE"]);
        const dateB = new Date(b["DATE"]);
        if (dateA - dateB !== 0) return dateA - dateB;
        return (a["From"] || "").localeCompare(b["From"] || "");
      });

      const groupKeys = [
        "SCHOOL",
        "TUTOR",
        "Session ID",
        "HRS",
        "DATE",
        "Duration",
        "Hourly Rate",
        "TOTAL $ P",
      ];

      const grouped = {};
      for (const row of data) {
        const key = groupKeys.map((k) => row[k]).join("|");
        if (!grouped[key]) grouped[key] = { ...row, NAME: new Set() };
        grouped[key].NAME.add(row["NAME"]);
      }

      let groupedArr = Object.values(grouped).map((obj) => ({
        ...obj,
        NAME: Array.from(obj.NAME).join("\n"),
      }));

      const cols = Object.keys(groupedArr[0]);
      const lastCol = cols.pop();
      cols.splice(2, 0, lastCol); // move NAME to position 2

      // --- Step 3: Draw paginated tables ---
      const pageWidth = 600;
      const pageHeight = 800;
      const margin = 20;
      const rowHeight = 18;
      const headerHeight = 25;
      const maxRowsPerPage = Math.floor((pageHeight - margin * 2 - headerHeight) / rowHeight);

      // Split groupedArr into chunks of maxRowsPerPage
      const pages = [];
      for (let i = 0; i < groupedArr.length; i += maxRowsPerPage) {
        pages.push(groupedArr.slice(i, i + maxRowsPerPage));
      }

      const pageCanvases = [];

      for (const [pageIndex, pageRows] of pages.entries()) {
        const canvas = document.createElement("canvas");
        canvas.width = pageWidth;
        canvas.height = pageHeight;
        const ctx = canvas.getContext("2d");

        ctx.fillStyle = "white";
        ctx.fillRect(0, 0, pageWidth, pageHeight);
        ctx.fillStyle = "black";
        ctx.font = "12px Arial";

        // Draw header
        let x = margin;
        let y = margin + 15;
        const colWidth = (pageWidth - margin * 2) / cols.length;
        ctx.font = "bold 12px Arial";
        cols.forEach((c, i) => {
          ctx.fillText(c, x + i * colWidth + 4, y);
        });

        ctx.font = "12px Arial";
        y += 10;

        // Draw rows
        for (const row of pageRows) {
          y += rowHeight;
          cols.forEach((c, i) => {
            const text = String(row[c] ?? "");
            const lines = text.split("\n");
            lines.forEach((line, j) => {
              ctx.fillText(line, x + i * colWidth + 4, y + j * 14);
            });
          });
        }

        pageCanvases.push(canvas);
      }

      // --- Step 4: Combine with first PDF page ---
      const pdfBytes = await pdfFile.arrayBuffer();
      const originalPdf = await PDFLib.PDFDocument.load(pdfBytes);
      const newPdf = await PDFLib.PDFDocument.create();

      const [firstPage] = await newPdf.copyPages(originalPdf, [0]);
      newPdf.addPage(firstPage);

      // Add each canvas page
      for (const canvas of pageCanvases) {
        const imgUrl = canvas.toDataURL("image/png");
        const img = await newPdf.embedPng(imgUrl);
        const page = newPdf.addPage([pageWidth, pageHeight]);
        page.drawImage(img, { x: 0, y: 0, width: pageWidth, height: pageHeight });
      }

      // --- Step 5: Save and download ---
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
});

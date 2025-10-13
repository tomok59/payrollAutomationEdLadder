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

      // --- Step 3: Draw paginated, styled tables (fixed newline/wrapping) ---
      const pageWidth = 595; // A4 width in pts
      const pageHeight = 842; // A4 height in pts
      const margin = 30;
      const headerHeight = 30;
      const fontFamily = "Arial";
      const fontSize = 10;
      const lineHeight = 14; // spacing between wrapped lines
      const padding = 6; // cell padding inside box

      // Compute column widths (evenly distributed for now)
      const colCount = cols.length;
      const colWidth = (pageWidth - margin * 2) / colCount;

      // Helper: wrap text but preserve existing newline characters
      function wrapTextPreserveNewlines(ctx, text, maxWidth) {
        if (text === null || text === undefined) return [""];
        // Split by explicit newlines first
        const paragraphs = String(text).split(/\r?\n/);
        const wrappedLines = [];
        for (const para of paragraphs) {
          // If paragraph empty, keep an empty line
          if (para.trim() === "") {
            wrappedLines.push("");
            continue;
          }
          const words = para.split(/\s+/);
          let line = "";
          for (const word of words) {
            const test = line ? line + " " + word : word;
            const width = ctx.measureText(test).width;
            if (width > maxWidth - padding * 2) {
              if (line) wrappedLines.push(line);
              // If single word is longer than the width, break the word
              if (ctx.measureText(word).width > maxWidth - padding * 2) {
                // break the long word into smaller chunks
                let chunk = "";
                for (const ch of word) {
                  const testChunk = chunk + ch;
                  if (ctx.measureText(testChunk).width > maxWidth - padding * 2) {
                    if (chunk) wrappedLines.push(chunk);
                    chunk = ch;
                  } else {
                    chunk = testChunk;
                  }
                }
                if (chunk) {
                  line = chunk;
                } else {
                  line = "";
                }
              } else {
                line = word;
              }
            } else {
              line = test;
            }
          }
          if (line) wrappedLines.push(line);
        }
        return wrappedLines;
      }

      // Build page chunks based on dynamic heights
      const pageCanvases = [];
      let currentRows = [];
      let currentHeight = margin + headerHeight + padding; // start height consumed by header

      // For text measurement we need a temporary canvas/context
      const measureCanvas = document.createElement("canvas");
      const measureCtx = measureCanvas.getContext("2d");
      measureCtx.font = `${fontSize}px ${fontFamily}`;

      const renderPage = (rows) => {
        const canvas = document.createElement("canvas");
        canvas.width = pageWidth;
        canvas.height = pageHeight;
        const ctx = canvas.getContext("2d");
        ctx.fillStyle = "white";
        ctx.fillRect(0, 0, pageWidth, pageHeight);

        // Header background
        ctx.fillStyle = "#555";
        ctx.fillRect(margin, margin, pageWidth - margin * 2, headerHeight);

        // Header text
        ctx.fillStyle = "white";
        ctx.font = `bold ${fontSize}px ${fontFamily}`;
        cols.forEach((c, i) => {
          const tx = margin + i * colWidth + padding;
          const ty = margin + Math.round(headerHeight / 2) + Math.round(fontSize / 2) - 2;
          ctx.fillText(c, tx, ty);
        });

        // Draw header grid lines
        ctx.strokeStyle = "#000";
        ctx.lineWidth = 0.5;
        for (let i = 0; i < cols.length; i++) {
          ctx.strokeRect(margin + i * colWidth, margin, colWidth, headerHeight);
        }

        // Rows
        let y = margin + headerHeight;
        ctx.font = `${fontSize}px ${fontFamily}`;
        ctx.fillStyle = "#000";

        for (const row of rows) {
          // Compute wrapped lines for each cell and rowHeight
          const cellLines = cols.map((c) =>
            wrapTextPreserveNewlines(measureCtx, String(row[c] ?? ""), colWidth)
          );
          const maxLines = Math.max(...cellLines.map((l) => (l.length === 0 ? 1 : l.length)));
          const rowHeight = maxLines * lineHeight + padding * 2;

          // Optional: alternating row background for readability
          // if desired, uncomment:
          // if ((rows.indexOf(row) % 2) === 1) {
          //   ctx.fillStyle = "#fafafa";
          //   ctx.fillRect(margin, y, pageWidth - margin*2, rowHeight);
          //   ctx.fillStyle = "#000";
          // }

          // Draw cell borders
          ctx.strokeStyle = "#999";
          ctx.lineWidth = 0.4;
          for (let i = 0; i < cols.length; i++) {
            ctx.strokeRect(margin + i * colWidth, y, colWidth, rowHeight);
          }

          // Draw text lines within each cell
          for (let i = 0; i < cols.length; i++) {
            const lines = cellLines[i].length ? cellLines[i] : [""];
            for (let j = 0; j < lines.length; j++) {
              const textLine = lines[j];
              const tx = margin + i * colWidth + padding;
              // Baseline: y + padding + (j * lineHeight) + font ascent approx
              const ty = y + padding + j * lineHeight + fontSize;
              ctx.fillText(textLine, tx, ty);
            }
          }

          y += rowHeight;
        }

        pageCanvases.push(canvas);
      };

      for (const row of groupedArr) {
        // For each new row we estimate its height using measurement context
        const cellLines = cols.map((c) =>
          wrapTextPreserveNewlines(measureCtx, String(row[c] ?? ""), colWidth)
        );
        const maxLines = Math.max(...cellLines.map((l) => (l.length === 0 ? 1 : l.length)));
        const rowHeight = maxLines * lineHeight + padding * 2;

        if (currentHeight + rowHeight > pageHeight - margin) {
          // render what we have and start a new page
          renderPage(currentRows);
          currentRows = [];
          currentHeight = margin + headerHeight + padding;
        }

        currentRows.push(row);
        currentHeight += rowHeight;
      }

      if (currentRows.length) renderPage(currentRows);

      // --- Step 4: Combine with original PDF ---
      const pdfBytes = await pdfFile.arrayBuffer();
      const originalPdf = await PDFLib.PDFDocument.load(pdfBytes);
      const newPdf = await PDFLib.PDFDocument.create();

      const [firstPage] = await newPdf.copyPages(originalPdf, [0]);
      newPdf.addPage(firstPage);

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
      alert("Error processing files. See console for details.");
    }
  });
});

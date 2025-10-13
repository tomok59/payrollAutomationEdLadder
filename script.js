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

      // --- Step 3: Page and style setup ---
      const pageWidth = 595;
      const pageHeight = 842;
      const margin = 30;
      const headerHeight = 30;
      const fontFamily = "Arial";
      const fontSize = 10;
      const lineHeight = 14;
      const padding = 6;

      const measureCanvas = document.createElement("canvas");
      const measureCtx = measureCanvas.getContext("2d");
      measureCtx.font = `${fontSize}px ${fontFamily}`;

      // --- Step 4: Calculate dynamic column widths ---
      const baseWidths = {};
      const minWidth = 60;
      const maxWidth = 180;

      cols.forEach((c) => {
        let maxText = c;
        for (const row of groupedArr.slice(0, 30)) {
          if (row[c] && String(row[c]).length > maxText.length) {
            maxText = String(row[c]);
          }
        }
        let w = measureCtx.measureText(maxText).width + padding * 2;
        if (c === "NAME") w *= 1.5; // give NAME extra space
        baseWidths[c] = Math.min(Math.max(w, minWidth), maxWidth);
      });

      const totalWidth = Object.values(baseWidths).reduce((a, b) => a + b, 0);
      const scale = (pageWidth - margin * 2) / totalWidth;
      const colWidths = cols.map((c) => baseWidths[c] * scale);

      // --- Step 5: Improved wrapping (no mid-name breaks) ---
      function wrapTextSmart(ctx, text, maxWidth, preserveNewlines = true) {
        if (!text) return [""];
        const paragraphs = preserveNewlines ? String(text).split(/\r?\n/) : [String(text)];
        const wrappedLines = [];

        for (const para of paragraphs) {
          const tokens = para.split(/(\s+|,)/).filter((t) => t.trim().length > 0 || t === " ");
          let line = "";
          for (const token of tokens) {
            const testLine = line ? line + token : token;
            const width = ctx.measureText(testLine).width;
            if (width > maxWidth - padding * 2) {
              if (line.trim()) wrappedLines.push(line.trim());
              line = token.trim();
            } else {
              line = testLine;
            }
          }
          if (line.trim()) wrappedLines.push(line.trim());
          else if (!para.trim()) wrappedLines.push("");
        }
        return wrappedLines.length ? wrappedLines : [""];
      }

      // --- Step 6: Render pages ---
      const pageCanvases = [];
      let currentRows = [];
      let currentHeight = margin + headerHeight + padding;

      const renderPage = (rows) => {
        const canvas = document.createElement("canvas");
        canvas.width = pageWidth;
        canvas.height = pageHeight;
        const ctx = canvas.getContext("2d");

        // Background
        ctx.fillStyle = "white";
        ctx.fillRect(0, 0, pageWidth, pageHeight);

        // Header
        ctx.fillStyle = "#555";
        ctx.fillRect(margin, margin, pageWidth - margin * 2, headerHeight);
        ctx.fillStyle = "white";
        ctx.font = `bold ${fontSize}px ${fontFamily}`;

        let x = margin;
        cols.forEach((c, i) => {
          const tx = x + padding;
          const ty = margin + 20;
          ctx.fillText(c, tx, ty);
          x += colWidths[i];
        });

        // Header grid
        ctx.strokeStyle = "#000";
        ctx.lineWidth = 0.5;
        x = margin;
        cols.forEach((c, i) => {
          ctx.strokeRect(x, margin, colWidths[i], headerHeight);
          x += colWidths[i];
        });

        // Rows
        let y = margin + headerHeight;
        ctx.font = `${fontSize}px ${fontFamily}`;
        ctx.fillStyle = "#000";

        for (const row of rows) {
          const cellLines = cols.map((c) =>
            wrapTextSmart(measureCtx, String(row[c] ?? ""), colWidths[cols.indexOf(c)])
          );
          const maxLines = Math.max(...cellLines.map((l) => l.length));
          const rowHeight = maxLines * lineHeight + padding * 2;

          // Optional: alternating row color
          if ((rows.indexOf(row) % 2) === 1) {
            ctx.fillStyle = "#f8f8f8";
            ctx.fillRect(margin, y, pageWidth - margin * 2, rowHeight);
            ctx.fillStyle = "#000";
          }

          // Borders and text
          x = margin;
          for (let i = 0; i < cols.length; i++) {
            ctx.strokeStyle = "#aaa";
            ctx.strokeRect(x, y, colWidths[i], rowHeight);
            const lines = cellLines[i];
            lines.forEach((line, j) => {
              const tx = x + padding;
              const ty = y + padding + j * lineHeight + fontSize;
              ctx.fillText(line, tx, ty);
            });
            x += colWidths[i];
          }

          y += rowHeight;
        }

        pageCanvases.push(canvas);
      };

      for (const row of groupedArr) {
        const cellLines = cols.map((c) =>
          wrapTextSmart(measureCtx, String(row[c] ?? ""), colWidths[cols.indexOf(c)])
        );
        const maxLines = Math.max(...cellLines.map((l) => l.length));
        const rowHeight = maxLines * lineHeight + padding * 2;

        if (currentHeight + rowHeight > pageHeight - margin) {
          renderPage(currentRows);
          currentRows = [];
          currentHeight = margin + headerHeight + padding;
        }

        currentRows.push(row);
        currentHeight += rowHeight;
      }

      if (currentRows.length) renderPage(currentRows);

      // --- Step 7: Combine with original PDF ---
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

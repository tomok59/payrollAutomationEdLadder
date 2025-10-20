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

      // --- Step 2: Read first page of PDF to get size ---
      const pdfBytes = await pdfFile.arrayBuffer();
      const tempPdf = await PDFLib.PDFDocument.load(pdfBytes);
      const [firstPage] = await tempPdf.copyPages(tempPdf, [0]);
      const { width: pageWidth, height: pageHeight } = firstPage.getSize();

      // --- Step 3: Define fixed column order ---
      const cols = [
        "SCHOOL",
        "TUTOR",
        "NAME",
        "Session ID",
        "HRS",
        "DATE",
        "Duration",
        "Hourly Rate",
        "TOTAL $ P",
      ];

      // --- Step 4: Group names for multiple entries ---
      const grouped = {};
      for (const row of data) {
        const key = cols.map((k) => row[k] ?? "").join("|");
        if (!grouped[key]) grouped[key] = { ...row, NAME: new Set() };
        grouped[key].NAME.add(row["NAME"]);
      }

      const groupedArr = Object.values(grouped).map((obj) => ({
        ...obj,
        NAME: Array.from(obj.NAME).join("\n"),
      }));

      // --- Step 5: Layout settings ---
      const margin = 30;
      const fontFamily = "Arial";
      const fontSize = 8;
      const lineHeight = 15;
      const padding = 8;

      const DPI = 300;
      const scaleFactor = DPI / 72;
      const renderWidth = Math.floor(pageWidth * scaleFactor);
      const renderHeight = Math.floor(pageHeight * scaleFactor);

      const measureCanvas = document.createElement("canvas");
      const measureCtx = measureCanvas.getContext("2d");
      measureCtx.font = `${fontSize}px ${fontFamily}`;

      // --- Step 6: Column widths ---
      const baseWidths = {};
      const minWidth = 60;
      const maxWidth = 200;

      cols.forEach((c) => {
        let widestText = c;
        for (const row of groupedArr.slice(0, 50)) {
          const val = row[c] ? String(row[c]) : "";
          if (val.length > widestText.length) widestText = val;
        }
        const headerWidth = measureCtx.measureText(c).width;
        const contentWidth = measureCtx.measureText(widestText).width;
        let w = Math.max(headerWidth, contentWidth) + padding * 2;

        if (c === "NAME") w *= 1.6;
        if (c === "DATE") w *= 1.6;
        if (c === "TOTAL $ P") w *= 1.1;
        if (c === "Hourly Rate") w *= 1.1;

        baseWidths[c] = Math.min(Math.max(w, minWidth), maxWidth);
      });

      const totalWidth = Object.values(baseWidths).reduce((a, b) => a + b, 0);
      const scale = (pageWidth - margin * 2) / totalWidth;
      const colWidths = cols.map((c) => baseWidths[c] * scale);

      function wrapTextSmart(ctx, text, maxWidth) {
        if (!text) return [""];
        const paragraphs = String(text).split(/\r?\n/);
        const linesOut = [];

        for (const para of paragraphs) {
          if (!para.trim()) {
            linesOut.push("");
            continue;
          }
          const tokens = para.split(/(\s+)/).filter((t) => t.length > 0);
          let line = "";
          for (const token of tokens) {
            const trial = line + token;
            const width = ctx.measureText(trial).width;
            if (width > maxWidth - padding * 2 && line.trim().length) {
              linesOut.push(line.trim());
              line = token.trimStart();
            } else {
              line = trial;
            }
          }
          if (line.trim().length) linesOut.push(line.trim());
        }
        return linesOut.length ? linesOut : [""];
      }

      // --- Step 7: Header height ---
      const headerLines = cols.map((c, i) =>
        wrapTextSmart(measureCtx, c, colWidths[i])
      );
      const maxHeaderLines = Math.max(...headerLines.map((l) => l.length));
      const headerHeight = maxHeaderLines * lineHeight + padding * 2;

      // --- Step 8: Render table pages ---
      const pageCanvases = [];
      let currentRows = [];
      let currentHeight = margin + headerHeight + padding;

      const renderPage = (rows) => {
        const canvas = document.createElement("canvas");
        canvas.width = renderWidth;
        canvas.height = renderHeight;
        const ctx = canvas.getContext("2d");
        ctx.scale(scaleFactor, scaleFactor);
        ctx.imageSmoothingQuality = "high";

        ctx.fillStyle = "white";
        ctx.fillRect(0, 0, pageWidth, pageHeight);

        // Header background
        ctx.fillStyle = "#2682e0"; // fixed blue color
        ctx.fillRect(margin, margin, pageWidth - margin * 2, headerHeight);

        // Header text
        ctx.fillStyle = "white";
        ctx.font = `bold ${fontSize}px ${fontFamily}`;
        let x = margin;
        cols.forEach((c, i) => {
          const lines = headerLines[i];
          const totalHeight = lines.length * lineHeight;
          let yStart = margin + (headerHeight - totalHeight) / 2 + fontSize / 2;

          for (let j = 0; j < lines.length; j++) {
            const tx = x + padding;
            const ty = yStart + j * lineHeight;
            ctx.fillText(lines[j], tx, ty);
          }
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

        // Table body
        let y = margin + headerHeight;
        ctx.font = `${fontSize}px ${fontFamily}`;
        ctx.fillStyle = "#000";

        for (const row of rows) {
          const cellLines = cols.map((c, i) =>
            wrapTextSmart(measureCtx, String(row[c] ?? ""), colWidths[i])
          );
          const maxLines = Math.max(...cellLines.map((l) => l.length));
          const rowHeight = maxLines * lineHeight + padding * 2;

          if ((rows.indexOf(row) % 2) === 1) {
            ctx.fillStyle = "#f5f5f5";
            ctx.fillRect(margin, y, pageWidth - margin * 2, rowHeight);
            ctx.fillStyle = "#000";
          }

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
        const cellLines = cols.map((c, i) =>
          wrapTextSmart(measureCtx, String(row[c] ?? ""), colWidths[i])
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

      // --- Step 9: Merge with original PDF ---
      const newPdf = await PDFLib.PDFDocument.create();
      const [origFirstPage] = await newPdf.copyPages(tempPdf, [0]);
      newPdf.addPage(origFirstPage);

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
      link.download = "invoice.pdf";
      link.textContent = "Download Invoice";
      link.classList.add("download-link");
      outputDiv.appendChild(link);
    } catch (err) {
      console.error(err);
      alert("Error processing files. See console for details.");
    }
  });
});

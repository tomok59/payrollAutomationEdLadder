document.getElementById("processBtn").addEventListener("click", async () => {
  const fileInput = document.getElementById("fileInput");
  const outputDiv = document.getElementById("output");
  outputDiv.innerHTML = "";

  if (!fileInput.files.length) {
    alert("Please select a PDF file first!");
    return;
  }

  const file = fileInput.files[0];
  const arrayBuffer = await file.arrayBuffer();

  try {
    const pdfDoc = await PDFLib.PDFDocument.load(arrayBuffer);
    const newPdf = await PDFLib.PDFDocument.create();
    const [firstPage] = await newPdf.copyPages(pdfDoc, [0]);
    newPdf.addPage(firstPage);

    const pdfBytes = await newPdf.save();
    const blob = new Blob([pdfBytes], { type: "application/pdf" });
    const url = URL.createObjectURL(blob);

    const link = document.createElement("a");
    link.href = url;
    link.download = file.name.replace(".pdf", "_first_page.pdf");
    link.textContent = "Download First Page PDF";
    link.classList.add("download-link");

    outputDiv.appendChild(link);
  } catch (error) {
    console.error(error);
    alert("There was an error processing the PDF.");
  }
});

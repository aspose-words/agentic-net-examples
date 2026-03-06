using Aspose.Words;
using Aspose.Words.Saving;

string inputPath = @"C:\Docs\Input.docx";          // Path to the source document
string outputPath = @"C:\Docs\OutputPdfA.pdf";    // Desired PDF/A output file

// Load the existing document
Document doc = new Document(inputPath);

// Create a save options object appropriate for PDF format
SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);

// Cast to PdfSaveOptions to access PDF‑specific settings
PdfSaveOptions pdfOptions = (PdfSaveOptions)saveOptions;

// Set the compliance level to PDF/A‑1b (you can choose other PDF/A levels as needed)
pdfOptions.Compliance = PdfCompliance.PdfA1b;

// Save the document as PDF/A
doc.Save(outputPath, pdfOptions);

using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source document.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Paths for the output files.
        string pdfAPath = @"C:\Docs\ResultPdfA.pdf";
        string pdfPath  = @"C:\Docs\ResultPdf.pdf";

        // Load the document.
        Document doc = new Document(inputPath);

        // ---------- Save as PDF/A ----------
        // Create a PdfSaveOptions instance via the factory method.
        PdfSaveOptions pdfAOptions = (PdfSaveOptions)SaveOptions.CreateSaveOptions(SaveFormat.Pdf);
        // Set the compliance level to PDF/A‑1b (you can choose any PDF/A level you need).
        pdfAOptions.Compliance = PdfCompliance.PdfA1b;
        // Save the document using the PDF/A options.
        doc.Save(pdfAPath, pdfAOptions);

        // ---------- Save as regular PDF ----------
        // Create default PDF save options (compliance defaults to PDF 1.7).
        PdfSaveOptions pdfOptions = (PdfSaveOptions)SaveOptions.CreateSaveOptions(SaveFormat.Pdf);
        // Save the document as a standard PDF.
        doc.Save(pdfPath, pdfOptions);
    }
}

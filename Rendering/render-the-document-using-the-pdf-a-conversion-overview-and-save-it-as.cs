using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source document (replace with actual path as needed)
        string sourcePath = @"MyDir\PdfAConversionOverview.docx";

        // Path where the PDF/A compliant PDF will be saved (replace with actual path as needed)
        string outputPath = @"ArtifactsDir\PdfAConversionOverview.pdf";

        // Load the existing Word document
        Document doc = new Document(sourcePath);

        // Create a SaveOptions object suitable for PDF format
        SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);

        // Cast to PdfSaveOptions to access PDF-specific properties
        PdfSaveOptions pdfOptions = (PdfSaveOptions)saveOptions;

        // Set the compliance level to PDF/A-1b (preserves visual appearance)
        pdfOptions.Compliance = PdfCompliance.PdfA1b;

        // Save the document as a PDF/A compliant PDF
        doc.Save(outputPath, pdfOptions);
    }
}

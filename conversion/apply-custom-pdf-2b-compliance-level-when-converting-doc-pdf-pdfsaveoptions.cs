using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfA2bConversion
{
    static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "ArtifactsDir");
        Directory.CreateDirectory(outputDir);

        // Output PDF file that will comply with PDF/A‑2b.
        // In Aspose.Words the PDF/A‑2b level is represented by PdfCompliance.PdfA2u.
        string outputPath = Path.Combine(outputDir, "Document.PdfA2b.pdf");

        // Create a new document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, PDF/A‑2b world!");

        // Create save options for PDF output.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            // Apply PDF/A‑2b compliance.
            Compliance = PdfCompliance.PdfA2u
        };

        // Save the document as PDF with the specified compliance level.
        doc.Save(outputPath, saveOptions);

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}

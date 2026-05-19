using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOC document in memory.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample DOC content for PDF/A‑2b conversion.");

        // Save the document locally as DOC (input file for the conversion).
        const string inputPath = "input.doc";
        source.Save(inputPath, SaveFormat.Doc);

        // Load the DOC document from the file system.
        Document doc = new Document(inputPath);

        // Configure PDF save options to use PDF/A‑2b compliance.
        // In Aspose.Words the PDF/A‑2b level is represented by PdfCompliance.PdfA2u.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u
        };

        // Save the document as PDF with the specified compliance level.
        const string outputPath = "output.pdf";
        doc.Save(outputPath, saveOptions);

        // Verify that the PDF file was created successfully.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected output PDF was not created.");
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names for the sample DOC and the resulting PDF/A‑2b file.
        const string inputPath = "sample.doc";
        const string outputPath = "sample_PdfA2b.pdf";

        // -----------------------------------------------------------------
        // Create a simple DOC document.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document for PDF/A‑2b conversion.");
        sourceDoc.Save(inputPath, SaveFormat.Doc);

        // -----------------------------------------------------------------
        // Load the DOC document.
        // -----------------------------------------------------------------
        Document doc = new Document(inputPath);

        // -----------------------------------------------------------------
        // Configure PDF save options for PDF/A‑2b compliance.
        // PDF/A‑2b is represented by the PdfA2u compliance level.
        // -----------------------------------------------------------------
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u
        };

        // -----------------------------------------------------------------
        // Save the document as PDF with the specified compliance level.
        // -----------------------------------------------------------------
        doc.Save(outputPath, saveOptions);

        // -----------------------------------------------------------------
        // Verify that the PDF file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"The expected PDF file '{outputPath}' was not created.");

        // Indicate success.
        Console.WriteLine("PDF/A‑2b conversion completed successfully.");
    }
}

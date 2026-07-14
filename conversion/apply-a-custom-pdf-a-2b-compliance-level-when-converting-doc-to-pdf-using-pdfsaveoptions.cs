using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ApplyPdfA2bCompliance
{
    public static void Main()
    {
        // Define file names.
        const string inputPath = "input.doc";
        const string outputPath = "output.pdf";

        // -----------------------------------------------------------------
        // Create a sample DOC file.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document for PDF/A‑2b conversion.");
        sourceDoc.Save(inputPath, SaveFormat.Doc);

        // -----------------------------------------------------------------
        // Load the DOC file.
        // -----------------------------------------------------------------
        Document doc = new Document(inputPath);

        // -----------------------------------------------------------------
        // Configure PDF save options for PDF/A‑2b compliance.
        // Aspose.Words does not expose a PdfA2b enum value; PdfA2u provides PDF/A‑2b
        // compliance with additional Unicode support, which satisfies the requirement.
        // -----------------------------------------------------------------
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u
        };

        // Save the document as PDF with the specified compliance level.
        doc.Save(outputPath, saveOptions);

        // -----------------------------------------------------------------
        // Verify that the PDF file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"The expected output file '{outputPath}' was not created.");

        Console.WriteLine("PDF/A‑2b conversion completed successfully.");
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string inputPath = "sample.doc";
        const string outputPath = "sample_pdfa1b.pdf";

        // -----------------------------------------------------------------
        // Create a simple DOC file.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample DOC file for PDF/A-1b conversion.");
        sourceDoc.Save(inputPath, SaveFormat.Doc);

        // -----------------------------------------------------------------
        // Load the DOC file.
        // -----------------------------------------------------------------
        Document doc = new Document(inputPath);

        // -----------------------------------------------------------------
        // Configure PDF save options for PDF/A-1b compliance.
        // -----------------------------------------------------------------
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };

        // -----------------------------------------------------------------
        // Save the document as PDF/A-1b.
        // -----------------------------------------------------------------
        doc.Save(outputPath, saveOptions);

        // -----------------------------------------------------------------
        // Verify that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"The expected output file '{outputPath}' was not created.");

        // Optional: indicate success (no interactive prompts required).
        Console.WriteLine("Conversion to PDF/A-1b completed successfully.");
    }
}

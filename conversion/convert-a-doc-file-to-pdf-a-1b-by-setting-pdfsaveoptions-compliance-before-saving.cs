using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOC file.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample DOC content for PDF/A-1b conversion.");
        string inputPath = "input.doc";
        source.Save(inputPath, SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document(inputPath);

        // Set PDF/A-1b compliance.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };

        // Save as PDF/A-1b.
        string outputPath = "output.pdf";
        doc.Save(outputPath, saveOptions);

        // Verify that the output file exists.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected output PDF/A-1b was not created.");
    }
}

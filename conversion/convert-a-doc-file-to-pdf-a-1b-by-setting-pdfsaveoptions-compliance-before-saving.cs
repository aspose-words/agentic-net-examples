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
        const string inputPath = "input.doc";
        source.Save(inputPath, SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document(inputPath);

        // Configure PDF save options for PDF/A-1b compliance.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };

        // Save the document as PDF/A-1b.
        const string outputPath = "output.pdf";
        doc.Save(outputPath, saveOptions);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected output PDF/A-1b was not created.");
    }
}

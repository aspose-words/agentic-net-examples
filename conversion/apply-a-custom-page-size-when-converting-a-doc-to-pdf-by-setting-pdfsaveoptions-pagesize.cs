using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple DOC file.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample DOC content.");
        const string inputPath = "input.doc";
        source.Save(inputPath, SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document(inputPath);

        // Set a custom page size (500x700 points) for the first section.
        // Points are the default unit for page dimensions in Aspose.Words.
        doc.FirstSection.PageSetup.PageWidth = 500f;
        doc.FirstSection.PageSetup.PageHeight = 700f;

        // Convert to PDF.
        const string outputPath = "output.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected output PDF was not created.");

        // Clean up temporary files (optional).
        File.Delete(inputPath);
    }
}

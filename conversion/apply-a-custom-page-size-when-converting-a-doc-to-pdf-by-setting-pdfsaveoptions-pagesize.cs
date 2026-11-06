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
        builder.Writeln("Sample DOC content with a custom PDF page size.");
        string inputPath = "input.doc";
        source.Save(inputPath, SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document(inputPath);

        // Set a custom page size (e.g., 5" x 7") for the first section.
        // 1 inch = 72 points, so 5" = 360 points, 7" = 504 points.
        Section firstSection = doc.FirstSection;
        firstSection.PageSetup.PageWidth = 360f;
        firstSection.PageSetup.PageHeight = 504f;

        // Save the document as PDF.
        string outputPath = "output.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected output PDF was not created.");

        // Optional: Clean up generated files (comment out if you want to inspect them).
        // File.Delete(inputPath);
        // File.Delete(outputPath);
    }
}

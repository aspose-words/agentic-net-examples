using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names
        const string inputPath = "sample.doc";
        const string outputPath = "custom_page_size.pdf";

        // Create a simple DOC file
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document with a custom PDF page size.");
        sourceDoc.Save(inputPath, SaveFormat.Doc);

        // Load the DOC file
        Document doc = new Document(inputPath);

        // Set a custom page size (A4: 595 x 842 points) for the first section
        Section firstSection = doc.FirstSection;
        firstSection.PageSetup.PaperSize = PaperSize.Custom;
        firstSection.PageSetup.PageWidth = 595f;   // Width in points
        firstSection.PageSetup.PageHeight = 842f;  // Height in points

        // Configure PDF save options (no need to set PageSize here)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save as PDF using the custom page size
        doc.Save(outputPath, pdfOptions);

        // Verify that the PDF was created
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF file was not created.");

        Console.WriteLine($"PDF successfully created at '{Path.GetFullPath(outputPath)}' with custom page size.");
    }
}

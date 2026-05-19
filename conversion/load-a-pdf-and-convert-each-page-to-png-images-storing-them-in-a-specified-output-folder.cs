using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Output directory for PDF and PNG files
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Path for the temporary PDF document
        string pdfPath = Path.Combine(outputDir, "sample.pdf");

        // Create a sample PDF with three pages
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");
        sampleDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the sample PDF.");

        // Load the PDF document
        Document pdfDoc = new Document(pdfPath);

        // Prepare image save options for PNG output
        ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            Resolution = 300 // optional: set DPI for higher quality
        };

        // Convert each page to a separate PNG file
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            imageOptions.PageSet = new PageSet(pageIndex);
            string pngPath = Path.Combine(outputDir, $"page_{pageIndex + 1}.png");
            pdfDoc.Save(pngPath, imageOptions);

            // Verify that the PNG file was created
            if (!File.Exists(pngPath))
                throw new InvalidOperationException($"Failed to create PNG for page {pageIndex + 1}.");
        }
    }
}

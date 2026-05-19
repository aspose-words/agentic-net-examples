using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("First page of the PDF.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page of the PDF.");
        sourceDoc.Save("sample.pdf", SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document("sample.pdf");

        // Configure image save options for PNG with 300 DPI resolution.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            Resolution = 300 // Sets both horizontal and vertical resolution.
        };

        // Export each page of the PDF to a separate PNG file.
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            pngOptions.PageSet = new PageSet(pageIndex);
            string outputFileName = $"page_{pageIndex + 1}.png";
            pdfDoc.Save(outputFileName, pngOptions);

            if (!File.Exists(outputFileName))
                throw new InvalidOperationException($"Failed to create PNG file: {outputFileName}");
        }

        // Clean up temporary PDF file.
        if (File.Exists("sample.pdf"))
            File.Delete("sample.pdf");
    }
}

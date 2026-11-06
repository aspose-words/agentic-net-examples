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
        builder.Writeln("This is the first page of the PDF.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is the second page of the PDF.");
        sourceDoc.Save("sample.pdf", SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document("sample.pdf");

        // Export each page to a PNG image with 300 DPI resolution.
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                Resolution = 300,
                PageSet = new PageSet(pageIndex)
            };

            string outputPath = $"page_{pageIndex + 1}.png";
            pdfDoc.Save(outputPath, pngOptions);

            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Expected output PNG '{outputPath}' was not created.");
        }
    }
}

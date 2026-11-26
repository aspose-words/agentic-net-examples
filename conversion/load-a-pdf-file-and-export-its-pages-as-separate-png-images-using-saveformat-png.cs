using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF file with two pages.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("First page content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page content.");
        sourceDoc.Save("sample.pdf", SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document("sample.pdf");

        // Export each page of the PDF as a separate PNG image.
        for (int i = 0; i < pdfDoc.PageCount; i++)
        {
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                PageSet = new PageSet(i) // Export only the current page.
            };

            string outputPath = $"page_{i + 1}.png";
            pdfDoc.Save(outputPath, options);

            // Verify that the PNG file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Expected PNG file '{outputPath}' was not created.");
        }
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample PDF file.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("This is page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 2.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 3.");

        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the sample PDF file.");

        // Step 2: Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Step 3: Export each page as a separate PNG image.
        for (int i = 0; i < pdfDoc.PageCount; i++)
        {
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render only the current page.
                PageSet = new PageSet(i)
            };

            string pngPath = $"page_{i + 1}.png";
            pdfDoc.Save(pngPath, options);

            if (!File.Exists(pngPath))
                throw new InvalidOperationException($"Failed to create PNG for page {i + 1}.");
        }

        // All pages have been exported successfully.
        Console.WriteLine("PDF pages have been exported to PNG images.");
    }
}

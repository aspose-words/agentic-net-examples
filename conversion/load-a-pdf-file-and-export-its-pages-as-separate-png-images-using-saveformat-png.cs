using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF file.
        const string pdfPath = "sample.pdf";
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Page 1 content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 content.");
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the sample PDF file.");

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Export each page as a separate PNG image.
        for (int i = 0; i < pdfDoc.PageCount; i++)
        {
            string pngPath = $"page_{i + 1}.png";

            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                PageSet = new PageSet(i) // Zero‑based page index.
            };

            pdfDoc.Save(pngPath, options);

            // Verify the PNG was created.
            if (!File.Exists(pngPath))
                throw new InvalidOperationException($"Failed to create PNG for page {i + 1}.");
        }

        // All pages have been exported successfully.
        Console.WriteLine("PDF pages have been exported to PNG images.");
    }
}

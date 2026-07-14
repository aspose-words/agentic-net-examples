using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF document.
        const string pdfPath = "sample.pdf";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 2.");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Load the PDF for conversion.
        Document pdfDoc = new Document(pdfPath);

        // Export each page to a separate PNG image with 300 DPI resolution.
        for (int i = 0; i < pdfDoc.PageCount; i++)
        {
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render only the current page.
                PageSet = new PageSet(i),
                // Set both horizontal and vertical resolution to 300 DPI.
                Resolution = 300
            };

            string pngPath = $"page_{i + 1}.png";
            pdfDoc.Save(pngPath, options);

            // Verify that the PNG image was created.
            if (!File.Exists(pngPath))
                throw new InvalidOperationException($"PNG file '{pngPath}' was not created.");
        }

        // All pages have been exported successfully.
        Console.WriteLine("PDF has been exported to PNG images at 300 DPI.");
    }
}

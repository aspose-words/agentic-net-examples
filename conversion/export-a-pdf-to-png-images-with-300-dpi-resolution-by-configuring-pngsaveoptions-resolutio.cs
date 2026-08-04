using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample document and save it as PDF (input.pdf).
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is page 1 of the PDF.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 2 of the PDF.");
        sourceDoc.Save("input.pdf", SaveFormat.Pdf);

        // Step 2: Load the PDF document we just created.
        Document pdfDoc = new Document("input.pdf");

        // Step 3: Export each page of the PDF to a separate PNG image at 300 DPI.
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                // Set both horizontal and vertical resolution to 300 DPI.
                Resolution = 300f,
                // Render only the current page.
                PageSet = new PageSet(pageIndex)
            };

            string outputFileName = $"page_{pageIndex + 1}.png";
            pdfDoc.Save(outputFileName, pngOptions);

            // Validate that the PNG file was created.
            if (!File.Exists(outputFileName))
                throw new InvalidOperationException($"Failed to create image file: {outputFileName}");
        }

        // All pages have been exported successfully.
        Console.WriteLine("PDF has been exported to PNG images at 300 DPI.");
    }
}

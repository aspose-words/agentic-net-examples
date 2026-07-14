using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample document with multiple pages.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Page 1 – Sample content for archival.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 – Additional content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 – Final page.");

        // Step 2: Save the document as PDF (the source for conversion).
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Step 3: Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Step 4: Render each page to a separate PNG image using lossless compression.
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            // Configure image save options for PNG.
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);
            // Render only the current page.
            pngOptions.PageSet = new PageSet(pageIndex);
            // Ensure no color reduction (PNG is lossless by default).
            pngOptions.ImageColorMode = ImageColorMode.None;
            // Optional: set a high resolution for archival quality.
            pngOptions.Resolution = 300;

            string pngFileName = $"page_{pageIndex + 1}.png";
            pdfDoc.Save(pngFileName, pngOptions);

            // Validate that the PNG file was written.
            if (!File.Exists(pngFileName))
                throw new InvalidOperationException($"Failed to create PNG for page {pageIndex + 1}.");
        }

        // All pages have been saved as PNG images.
        Console.WriteLine("PDF successfully converted to PNG image sequence.");
    }
}

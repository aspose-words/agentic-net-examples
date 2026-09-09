using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing; // For BreakType

public class PdfToPngConverter
{
    public static void Main()
    {
        // Define file names
        const string pdfPath = "sample.pdf";

        // -----------------------------------------------------------------
        // Step 1: Create a sample PDF document with multiple pages.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Add content for three pages
        builder.Writeln("This is page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 2.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 3.");

        // Save the document as PDF
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify PDF creation
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the source PDF file.");

        // -----------------------------------------------------------------
        // Step 2: Load the PDF and export each page as a separate PNG.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // Iterate through all pages in the PDF
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            // Configure image save options for PNG format
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render only the current page
                PageSet = new PageSet(pageIndex),

                // Set resolution (dpi). ImageSize is optional and omitted to avoid type conflicts.
                Resolution = 300
            };

            // Define output PNG file name
            string pngPath = $"page_{pageIndex + 1}.png";

            // Save the current page as PNG
            pdfDoc.Save(pngPath, pngOptions);

            // Validate that the PNG file was created
            if (!File.Exists(pngPath) || new FileInfo(pngPath).Length == 0)
                throw new InvalidOperationException($"Failed to create PNG for page {pageIndex + 1}.");
        }

        // All pages have been exported successfully.
        Console.WriteLine("PDF pages have been exported to PNG images.");
    }
}

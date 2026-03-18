using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToPngExport
{
    static void Main()
    {
        // Load the source PDF document.
        Document pdfDoc = new Document("Input.pdf");

        // Configure image save options for PNG format with 300 DPI resolution.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);
        pngOptions.Resolution = 300; // Sets both horizontal and vertical DPI.

        // Export each page of the PDF as a separate PNG image.
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            // Render only the current page.
            pngOptions.PageSet = new PageSet(pageIndex);

            // Save the page to a PNG file.
            string outputPath = $"Page_{pageIndex + 1}.png";
            pdfDoc.Save(outputPath, pngOptions);
        }
    }
}

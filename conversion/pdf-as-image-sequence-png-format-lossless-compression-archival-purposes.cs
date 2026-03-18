using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToPngSequence
{
    static void Main()
    {
        // Load the source PDF document.
        Document pdfDoc = new Document("input.pdf");

        // Iterate through each page of the PDF.
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            // Create ImageSaveOptions for PNG format (lossless by definition).
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);

            // Render only the current page.
            pngOptions.PageSet = new PageSet(pageIndex);

            // Save the rendered page as a PNG file.
            string outputPath = $"output_page_{pageIndex + 1}.png";
            pdfDoc.Save(outputPath, pngOptions);
        }
    }
}

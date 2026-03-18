using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class PdfHyphenationToImage
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfFile = @"C:\Docs\Source.pdf";

        // Path where the rendered image will be saved.
        string imageFile = @"C:\Docs\Rendered.png";

        // Load the PDF into a Document object using PdfLoadOptions.
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        Document doc = new Document(pdfFile, loadOptions);

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Optional: adjust hyphenation parameters if desired.
        // doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        // doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch

        // Prepare image save options.
        ImageSaveOptions imgOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Render the first page only. Remove or modify to render all pages.
            PageSet = new PageSet(0),

            // Set a reasonable resolution (dpi) for visual inspection.
            Resolution = 150
        };

        // Save the rendered page to an image file.
        doc.Save(imageFile, imgOptions);
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertDocumentToImages
{
    static void Main()
    {
        // Load the source multi‑page document (any supported format).
        Document doc = new Document("InputDocument.docx");

        // Create ImageSaveOptions to render pages as PNG images.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Render all pages of the document.
            PageSet = PageSet.All,

            // Optional: set resolution (dpi) for higher quality images.
            HorizontalResolution = 300,
            VerticalResolution = 300,

            // Optional: keep the default color mode (full color).
            ImageColorMode = ImageColorMode.None
        };

        // Save each page as a separate PNG file.
        // Aspose.Words will append the page index to the file name.
        doc.Save("OutputPage.png", saveOptions);
    }
}

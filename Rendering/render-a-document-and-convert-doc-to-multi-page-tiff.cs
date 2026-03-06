using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("Input.doc");

        // Set up image save options for TIFF output.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // Render each page as a separate frame in a single multi‑page TIFF file.
        saveOptions.PageLayout = MultiPageLayout.TiffFrames();
        // Optional: define the resolution (dots per inch) of the rendered images.
        saveOptions.Resolution = 300;

        // Save the document as a multi‑page TIFF.
        doc.Save("Output.tiff", saveOptions);
    }
}

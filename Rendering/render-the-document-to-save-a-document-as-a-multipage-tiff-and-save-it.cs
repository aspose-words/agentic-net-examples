using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("input.docx");

        // Create ImageSaveOptions for TIFF format.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);

        // Render each page as a separate frame in a multi‑page TIFF.
        tiffOptions.PageLayout = MultiPageLayout.TiffFrames();

        // Optional: set compression and resolution.
        tiffOptions.TiffCompression = TiffCompression.Lzw;
        tiffOptions.Resolution = 300;

        // Save the document as a multipage TIFF file.
        doc.Save("output.tiff", tiffOptions);
    }
}

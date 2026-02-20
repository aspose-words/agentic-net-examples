using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("Input.doc");

        // Set up options to save each page as a separate frame in a multi‑page TIFF.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.PageLayout = MultiPageLayout.TiffFrames();   // multi‑page TIFF layout
        options.TiffCompression = TiffCompression.Lzw;      // optional compression
        options.Resolution = 300;                           // optional DPI setting

        // Save the document as a multi‑page TIFF file.
        doc.Save("Output.tiff", options);
    }
}

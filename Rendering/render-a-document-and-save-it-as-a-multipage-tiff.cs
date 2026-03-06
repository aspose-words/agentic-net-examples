using Aspose.Words;
using Aspose.Words.Saving;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("Input.docx");

        // Create image save options for TIFF format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);

        // Set the layout so that each page is saved as a separate frame in the TIFF.
        options.PageLayout = MultiPageLayout.TiffFrames();

        // Optional: adjust resolution and compression.
        options.Resolution = 300;               // DPI
        options.TiffCompression = TiffCompression.Lzw;

        // Save the document as a multipage TIFF file.
        doc.Save("Output.tiff", options);
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("InputDocument.docx");

        // Configure image save options for TIFF format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);

        // Use the TiffFrames layout so each page becomes a separate frame in a multipage TIFF.
        options.PageLayout = MultiPageLayout.TiffFrames();

        // Optional: set compression (default is Lzw). Uncomment to change.
        // options.TiffCompression = TiffCompression.Lzw;

        // Save the document as a multipage TIFF file.
        doc.Save("OutputDocument.tiff", options);
    }
}

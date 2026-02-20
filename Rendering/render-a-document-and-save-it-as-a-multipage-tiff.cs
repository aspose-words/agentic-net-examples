using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Set up options for rendering the document to a multipage TIFF.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        // Each page will be saved as a separate frame in the TIFF file.
        options.PageLayout = MultiPageLayout.TiffFrames();
        // Example compression; can be changed as needed.
        options.TiffCompression = TiffCompression.Lzw;

        // Save the document as a multipage TIFF.
        doc.Save("Output.tiff", options);
    }
}

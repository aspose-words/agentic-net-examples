using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the multi‑page document from a file.
        Document doc = new Document("input.docx");

        // Configure image save options to render all pages into a single multi‑frame TIFF image.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        // Use the TiffFrames layout so each page becomes a separate frame in the TIFF.
        options.PageLayout = MultiPageLayout.TiffFrames();

        // Save the document as an image. This will convert every page of the document
        // into a frame of the resulting TIFF file.
        doc.Save("output.tiff", options);
    }
}

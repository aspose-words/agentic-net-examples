using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing Word document.
        // Replace "input.docx" with the path to your source document.
        Document doc = new Document("input.docx");

        // Create ImageSaveOptions for TIFF format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);

        // Use the default multi‑page layout for TIFF (each page as a separate frame).
        // This is equivalent to MultiPageLayout.TiffFrames().
        options.PageLayout = MultiPageLayout.TiffFrames();

        // Optional: set compression type (LZW is the default).
        // options.TiffCompression = TiffCompression.Lzw;

        // Save the document as a multipage TIFF.
        // Replace "output.tiff" with the desired output file path.
        doc.Save("output.tiff", options);
    }
}

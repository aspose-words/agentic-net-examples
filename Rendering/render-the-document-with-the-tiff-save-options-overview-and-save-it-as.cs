using Aspose.Words;
using Aspose.Words.Saving;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample content to the document.
        builder.Writeln("This is a sample document rendered to TIFF.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page of the document.");

        // Create ImageSaveOptions for TIFF format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);

        // Set the desired TIFF compression (LZW is the default).
        options.TiffCompression = TiffCompression.Lzw;

        // Optional: adjust resolution and image size.
        options.Resolution = 300; // 300 DPI
        options.ImageSize = new Size(2480, 3508); // A4 size at 300 DPI

        // Save the document as a TIFF image using the specified options.
        string outputPath = "OutputDocument.tiff";
        doc.Save(outputPath, options);
    }
}

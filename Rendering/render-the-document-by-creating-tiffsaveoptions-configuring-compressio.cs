using Aspose.Words;
using Aspose.Words.Saving;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample content spanning three pages.
        builder.Writeln("First page.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Third page.");

        // Configure image save options for TIFF output.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        // Set the desired TIFF compression.
        options.TiffCompression = TiffCompression.Lzw;
        // Set the resolution (dots per inch) for both axes.
        options.Resolution = 300;
        // Define the page range to render (pages 2 and 3, zero‑based indices).
        options.PageSet = new PageSet(new PageRange(1, 2));

        // Save the document as a multi‑page TIFF file.
        doc.Save("RenderedDocument.tiff", options);
    }
}

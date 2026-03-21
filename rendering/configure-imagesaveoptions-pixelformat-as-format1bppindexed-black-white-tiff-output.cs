using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a simple Word document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, world!");

        // Configure image save options for 1‑bpp black‑and‑white TIFF output.
        var options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            PixelFormat = ImagePixelFormat.Format1bppIndexed,
            TiffCompression = TiffCompression.Ccitt4,
            ImageColorMode = ImageColorMode.BlackAndWhite
        };

        // Save the document as a black‑and‑white TIFF file.
        doc.Save("Output.tiff", options);
    }
}

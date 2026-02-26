using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsTiffExample
{
    class Program
    {
        static void Main()
        {
            // Load the source DOC document from the file system.
            // The Document constructor automatically detects the format.
            Document doc = new Document("InputDocument.doc");

            // Create ImageSaveOptions specifying the TIFF format.
            // This object allows us to control rendering options such as resolution or compression.
            ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);

            // Example: set the resolution to 300 DPI for higher quality output.
            tiffOptions.Resolution = 300;

            // Example: use LZW compression (default) – can be changed if needed.
            // tiffOptions.TiffCompression = TiffCompression.Lzw;

            // Save the rendered document as a TIFF file.
            // The Save method with (string, SaveOptions) follows the provided lifecycle rule.
            doc.Save("OutputDocument.tiff", tiffOptions);
        }
    }
}

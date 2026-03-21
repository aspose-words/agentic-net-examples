using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document with some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for TIFF rendering.");

        // Disable OpenType font formatting features to ensure minimal rendering size.
        doc.CompatibilityOptions.DisableOpenTypeFontFormattingFeatures = true;

        // Configure image save options for 1bpp TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT Group 4 compression (1-bit per pixel).
            TiffCompression = TiffCompression.Ccitt4,
            // Force black-and-white output.
            ImageColorMode = ImageColorMode.BlackAndWhite,
            // Set resolution if needed.
            Resolution = 300
        };

        // Save the document as a TIFF image.
        doc.Save("output.tif", saveOptions);
    }
}

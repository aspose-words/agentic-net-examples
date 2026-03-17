using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("input.docx");

        // Disable OpenType font formatting features to ensure minimal rendering size.
        // This uses the CompatibilityOptions.DisableOpenTypeFontFormattingFeatures property.
        doc.CompatibilityOptions.DisableOpenTypeFontFormattingFeatures = true;

        // Configure image save options for 1bpp TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT Group 4 compression (1-bit per pixel).
            TiffCompression = TiffCompression.Ccitt4,
            // Force black-and-white output.
            ImageColorMode = ImageColorMode.BlackAndWhite,
            // Optional: set resolution if needed.
            Resolution = 300
        };

        // Save the document as a TIFF image.
        doc.Save("output.tif", saveOptions);
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Saving;
using System.Drawing;

class GrayscaleTiffExport
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample content.
        builder.ParagraphFormat.Style = doc.Styles["Heading 1"];
        builder.Writeln("Archive Document");
        builder.Writeln("This document will be saved as a grayscale TIFF for efficient archiving.");

        // Insert an image to demonstrate rendering.
        // (Replace with a valid path to an image file on your system.)
        // builder.InsertImage(@"C:\Images\Sample.jpg");

        // Configure image save options for TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render pages in grayscale.
            ImageColorMode = ImageColorMode.Grayscale,

            // Use LZW compression to keep file size reasonable while preserving quality.
            TiffCompression = TiffCompression.Lzw,

            // Optional: set a higher resolution for better detail (e.g., 300 DPI).
            Resolution = 300,

            // Optional: set a white paper background (default is white).
            PaperColor = Color.White
        };

        // Save the document as a grayscale TIFF.
        doc.Save("ArchiveDocument.tiff", saveOptions);
    }
}

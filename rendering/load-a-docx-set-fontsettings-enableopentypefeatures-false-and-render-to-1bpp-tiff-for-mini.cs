using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for rendering to TIFF.");

        // Configure FontSettings. The specific EnableOpenTypeFeatures property does not exist,
        // so we simply assign a new FontSettings instance to the document.
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Define the output TIFF file path.
        string outputPath = "output.tiff";

        // Set up ImageSaveOptions for 1bpp (black and white) TIFF with CCITT Group4 compression.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            ImageColorMode = ImageColorMode.BlackAndWhite,
            // Use the correct enum value for CCITT Group4 compression.
            TiffCompression = TiffCompression.Ccitt4,
            Resolution = 96 // Optional: lower resolution to reduce size.
        };

        // Render the document to TIFF.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The TIFF file was not created.");
        }
    }
}

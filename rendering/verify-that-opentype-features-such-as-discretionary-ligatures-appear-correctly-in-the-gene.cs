using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Path for the rendered TIFF file.
        string tiffPath = "DiscretionaryLigature.tiff";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a common font that contains ligature glyphs.
        builder.Font.Name = "Arial";
        builder.Font.Size = 48;

        // Insert text that includes the discretionary ligature character (ﬁ – U+FB01).
        // This character is a precomposed ligature glyph, ensuring it will be rendered
        // if the font supports it, without needing to enable OpenType features via API.
        builder.Writeln("Discretionary ligature test: ﬁ");

        // Configure image save options for TIFF output.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Optional: increase resolution for clearer rendering.
            Resolution = 300
        };

        // Render the document to a single-page TIFF image.
        doc.Save(tiffPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException($"Failed to create TIFF file at '{tiffPath}'.");

        // The example finishes without requiring user interaction.
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a font that supports OpenType ligatures (e.g., "Times New Roman").
        builder.Font.Name = "Times New Roman";
        builder.Font.Size = 48;

        // Text containing ligatures (ff, fi, fl) and a stylistic set example.
        builder.Writeln("Office: office, efficient, flake.");

        // Ensure OpenType features are not disabled via compatibility options.
        doc.CompatibilityOptions.DisableOpenTypeFontFormattingFeatures = false;

        // Assign default font settings (optional, demonstrates explicit configuration).
        doc.FontSettings = new FontSettings();

        // Configure image save options for TIFF output.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render with high quality to preserve typographic features.
            UseAntiAliasing = true,
            UseHighQualityRendering = true,
            // Set a reasonable resolution.
            Resolution = 300
        };

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RenderedOutput.tiff");

        // Save the document as a TIFF image.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the TIFF output file.");

        // Optionally, report success.
        Console.WriteLine($"TIFF rendered successfully: {outputPath}");
    }
}

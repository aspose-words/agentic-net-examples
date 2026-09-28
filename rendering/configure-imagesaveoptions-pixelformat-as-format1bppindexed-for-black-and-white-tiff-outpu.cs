using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document with one paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for black‑and‑white TIFF conversion.");

        // Define the output file path.
        string outputPath = "output.tiff";

        // Configure ImageSaveOptions for TIFF with black‑and‑white (1‑bit per pixel) output.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // Use ImageColorMode.BlackAndWhite to produce a 1‑bpp indexed image.
        saveOptions.ImageColorMode = ImageColorMode.BlackAndWhite;

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Indicate success.
        Console.WriteLine("TIFF file saved successfully: " + Path.GetFullPath(outputPath));
    }
}

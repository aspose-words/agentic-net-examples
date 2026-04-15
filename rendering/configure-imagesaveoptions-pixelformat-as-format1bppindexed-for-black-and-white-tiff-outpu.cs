using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document with a heading and a paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Black‑and‑White TIFF Example");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This page will be saved as a 1‑bpp indexed TIFF image.");

        // Prepare the output folder.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "BlackWhite.tiff");

        // Configure ImageSaveOptions for TIFF with 1‑bpp indexed pixel format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            PixelFormat = ImagePixelFormat.Format1bppIndexed,
            // Optional: use a CCITT compression scheme suitable for 1‑bpp images.
            TiffCompression = TiffCompression.Ccitt4
        };

        // Save the document as a TIFF image.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The TIFF file was not created.", outputPath);

        // Optionally, report success (no console input required).
        Console.WriteLine("TIFF image saved successfully to: " + outputPath);
    }
}

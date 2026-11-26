using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document with some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, world! This is a test for 1‑bit black‑white TIFF output.");

        // Define the output file path.
        string outputPath = "output_1bpp.tiff";

        // Configure ImageSaveOptions for TIFF with 1‑bit indexed pixel format.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        saveOptions.PixelFormat = ImagePixelFormat.Format1bppIndexed;

        // Save the document as a TIFF image.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The TIFF file was not created.");
        }

        // Indicate successful completion.
        Console.WriteLine("TIFF saved successfully: " + Path.GetFullPath(outputPath));
    }
}

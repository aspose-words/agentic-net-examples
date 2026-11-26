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
        builder.Writeln("Hello World!");

        // Set up image save options for black‑and‑white TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render the image in black‑and‑white mode.
            ImageColorMode = ImageColorMode.BlackAndWhite,
            // Use 1‑bit indexed pixel format for true B&W output.
            PixelFormat = ImagePixelFormat.Format1bppIndexed
        };

        // Save the document as a TIFF file.
        string outputFile = "output.tiff";
        doc.Save(outputFile, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        Console.WriteLine("TIFF saved successfully: " + Path.GetFullPath(outputFile));
    }
}

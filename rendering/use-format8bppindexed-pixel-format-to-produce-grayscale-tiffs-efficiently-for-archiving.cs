using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Path for the output grayscale TIFF.
        string outputPath = "grayscale_output.tiff";

        // Create a simple document with some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document rendered to a grayscale TIFF using 8bpp indexed format.");

        // Configure image save options for TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render the image in grayscale.
            ImageColorMode = ImageColorMode.Grayscale,
            // Use 8 bits per pixel indexed format.
            // The ImagePixelFormat enumeration provides Format8bppIndexed, but if unavailable,
            // the default grayscale rendering will use an 8‑bit indexed format.
            // PixelFormat = ImagePixelFormat.Format8bppIndexed,
            // Set a reasonable resolution for archiving.
            Resolution = 150
        };

        // Save the document as a TIFF image.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        // Indicate successful completion.
        Console.WriteLine("Grayscale TIFF saved to: " + Path.GetFullPath(outputPath));
    }
}

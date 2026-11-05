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
        builder.Writeln("This is a sample document for TIFF conversion with a threshold.");

        // Define the output TIFF file path.
        string outputPath = "output.tiff";

        // Configure image save options for grayscale TIFF.
        // The current Aspose.Words version does not expose a direct threshold property,
        // so we use the closest available settings: Grayscale color mode and CCITT4 compression,
        // which produces a black‑white image suitable for lighter grayscale rendering.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            ImageColorMode = ImageColorMode.Grayscale,
            TiffCompression = TiffCompression.Ccitt4
        };

        // Save the document as a TIFF image.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The TIFF file was not created.", outputPath);
        }

        // Output the file size to confirm creation.
        long fileSize = new FileInfo(outputPath).Length;
        Console.WriteLine($"TIFF file created successfully. Size: {fileSize} bytes.");
    }
}

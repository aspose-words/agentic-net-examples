using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document with a paragraph of text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document to be saved as a TIFF image.");

        // Configure ImageSaveOptions for TIFF output.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Adjust the dithering threshold for moderately dark images.
            ThresholdForFloydSteinbergDithering = 180
        };

        // Define the output file path.
        string outputPath = "output.tiff";

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The TIFF file was not created.", outputPath);
        }

        // Optionally, output a confirmation message.
        Console.WriteLine($"TIFF image saved successfully to '{Path.GetFullPath(outputPath)}'.");
    }
}

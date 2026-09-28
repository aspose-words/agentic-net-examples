using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document with some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for TIFF rendering with a higher dithering threshold.");

        // Configure ImageSaveOptions for TIFF and set the dithering threshold to 200.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        saveOptions.ThresholdForFloydSteinbergDithering = 200;

        string outputFile = "output.tiff";

        // Render the document to a TIFF file using the configured options.
        doc.Save(outputFile, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputFile))
        {
            throw new InvalidOperationException("Failed to create the TIFF file.");
        }

        // Output the file size as a simple confirmation.
        long fileSize = new FileInfo(outputFile).Length;
        Console.WriteLine($"TIFF saved successfully. Size: {fileSize} bytes.");
    }
}

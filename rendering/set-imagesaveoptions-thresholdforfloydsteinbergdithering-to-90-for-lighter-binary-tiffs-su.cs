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
        builder.Writeln("Sample text for OCR preprocessing.");

        // Configure TIFF save options with a high dithering threshold.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        saveOptions.ThresholdForFloydSteinbergDithering = 90;
        // Ensure the output is a binary (black‑and‑white) image.
        saveOptions.ImageColorMode = ImageColorMode.BlackAndWhite;

        string outputFile = "Sample.tiff";

        // Render the document to a TIFF file.
        doc.Save(outputFile, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputFile))
        {
            throw new InvalidOperationException($"Failed to create TIFF file at '{outputFile}'.");
        }

        // Output basic information about the saved file.
        long fileSize = new FileInfo(outputFile).Length;
        Console.WriteLine($"TIFF saved successfully. Size: {fileSize} bytes.");
    }
}

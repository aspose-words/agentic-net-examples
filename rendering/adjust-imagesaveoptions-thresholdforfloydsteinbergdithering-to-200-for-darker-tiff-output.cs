using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading and some text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Sample Document");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This document is rendered to a TIFF image with a higher dithering threshold for darker output.");

        // Configure TIFF save options with Floyd‑Steinberg dithering and a threshold of 200.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt3,
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            ThresholdForFloydSteinbergDithering = 200
        };

        // Save the document as a TIFF file.
        string outputPath = Path.Combine(artifactsDir, "DitheredOutput.tiff");
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The TIFF file was not created.", outputPath);
    }
}

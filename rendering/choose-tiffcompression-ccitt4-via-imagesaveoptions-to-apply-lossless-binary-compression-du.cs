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
        builder.Writeln("Hello, Aspose.Words TIFF rendering with CCITT4 compression!");

        // Configure TIFF save options to use CCITT4 (lossless binary) compression.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt4
        };

        // Define output file path.
        string outputPath = "output.tiff";

        // Save the document as a TIFF image using the specified options.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("TIFF file was not created.", outputPath);
        }

        // Optionally, output the file size to confirm saving succeeded.
        long fileSize = new FileInfo(outputPath).Length;
        Console.WriteLine($"TIFF file saved successfully: {outputPath} ({fileSize} bytes)");
    }
}

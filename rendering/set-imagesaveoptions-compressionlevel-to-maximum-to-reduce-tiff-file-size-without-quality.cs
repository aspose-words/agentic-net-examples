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
        builder.Writeln("Sample text for TIFF rendering with maximum compression.");

        // Configure ImageSaveOptions for TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // Use LZW compression (lossless) to achieve maximum compression for TIFF.
        saveOptions.TiffCompression = TiffCompression.Lzw;

        // Save the document as a TIFF file.
        string outputPath = "output.tiff";
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the TIFF file.");

        // Output the size of the generated TIFF file.
        long fileSize = new FileInfo(outputPath).Length;
        Console.WriteLine($"TIFF file saved successfully. Size: {fileSize} bytes.");
    }
}

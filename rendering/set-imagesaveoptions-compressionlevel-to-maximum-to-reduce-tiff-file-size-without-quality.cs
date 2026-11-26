using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document to demonstrate TIFF compression.");

        // Insert a shape to increase the amount of content.
        builder.InsertShape(ShapeType.Rectangle, 300, 200);

        // Set up ImageSaveOptions for TIFF with maximum lossless compression (LZW).
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        saveOptions.TiffCompression = TiffCompression.Lzw; // LZW provides lossless compression.

        string outputPath = "output.tiff";

        // Save the document as a TIFF image.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        // Output the size of the generated file.
        long fileSize = new FileInfo(outputPath).Length;
        Console.WriteLine($"TIFF saved successfully. Size: {fileSize} bytes.");
    }
}

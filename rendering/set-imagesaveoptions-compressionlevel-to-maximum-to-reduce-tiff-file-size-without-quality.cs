using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample content.
        builder.Writeln("This is a sample document.");
        builder.Writeln("It will be saved as a TIFF image with maximum loss‑less compression.");

        // Configure image save options for TIFF.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        // Use LZW compression, which is lossless and provides the highest compression among the supported schemes.
        options.TiffCompression = TiffCompression.Lzw;

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.tiff");

        // Save the document as a TIFF image.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The TIFF file was not created.", outputPath);

        // Optionally, report the file size.
        long fileSize = new FileInfo(outputPath).Length;
        Console.WriteLine($"TIFF file saved successfully: {outputPath}");
        Console.WriteLine($"File size: {fileSize} bytes");
    }
}

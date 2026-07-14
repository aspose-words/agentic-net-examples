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
        builder.Writeln("This is a sample document rendered to a TIFF image.");
        builder.Writeln("The image will be saved using CCITT4 compression.");

        // Configure image save options for TIFF with CCITT4 compression.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt4
        };

        // Define the output file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "SampleCcitt4.tiff");

        // Ensure the directory exists.
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

        // Save the document as a TIFF image using the specified options.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the TIFF file at '{outputPath}'.");
        }

        // Optionally, output the file size for confirmation.
        long fileSize = new FileInfo(outputPath).Length;
        Console.WriteLine($"TIFF file saved successfully: {outputPath}");
        Console.WriteLine($"File size: {fileSize} bytes");
    }
}

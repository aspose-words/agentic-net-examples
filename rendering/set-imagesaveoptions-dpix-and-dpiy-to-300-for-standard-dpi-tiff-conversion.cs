using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document for TIFF conversion with 300 DPI.");

        // Configure image save options for TIFF with 300 DPI.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // Set horizontal and vertical resolution (DPI) to 300.
        saveOptions.HorizontalResolution = 300f;
        saveOptions.VerticalResolution = 300f;

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Define the output file path.
        string outputPath = Path.Combine(outputDir, "Sample_300dpi.tiff");

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("TIFF file was not created.");

        // Optionally, output the file size for confirmation.
        long fileSize = new FileInfo(outputPath).Length;
        Console.WriteLine($"TIFF file saved successfully: {outputPath}");
        Console.WriteLine($"File size: {fileSize} bytes");
    }
}

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
        builder.Writeln("This is a sample document rendered to a high‑resolution TIFF image.");

        // Configure image save options for TIFF with a resolution of 300 DPI.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // The property to set DPI is Resolution, not DesiredDpi.
            Resolution = 300
        };

        // Define the output file path.
        string outputPath = "output.tiff";

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The TIFF image was not created.", outputPath);
        }

        // Output the file size to confirm that the image was saved.
        long fileSize = new FileInfo(outputPath).Length;
        Console.WriteLine($"TIFF image saved successfully. File size: {fileSize} bytes.");
    }
}

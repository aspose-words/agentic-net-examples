using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Output file path for the low‑resolution TIFF.
        const string outputPath = "LowResolution.tiff";

        // Create a simple document with some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a low‑resolution TIFF suitable for email attachments.");

        // Configure image save options for TIFF.
        // The DesiredDpi property does not exist; use the Resolution property instead.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Set both horizontal and vertical DPI to 72.
            Resolution = 72
        };

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Indicate success.
        Console.WriteLine("TIFF file created successfully at: " + Path.GetFullPath(outputPath));
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, TIFF!");

        // Define the output TIFF file path.
        string tiffPath = "output.tiff";

        // Configure TIFF save options with CCITT Group 4 compression.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT Group 4 compression (enum value is Ccitt4).
            TiffCompression = TiffCompression.Ccitt4
        };

        // Render the document to TIFF.
        doc.Save(tiffPath, tiffOptions);

        // Verify that the TIFF file was created and is not empty.
        if (!File.Exists(tiffPath))
            throw new Exception("TIFF file was not created.");

        FileInfo fileInfo = new FileInfo(tiffPath);
        if (fileInfo.Length == 0)
            throw new Exception("TIFF file is empty.");

        Console.WriteLine("TIFF output verification passed.");
    }
}

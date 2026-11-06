using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a TrueType font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("This is a sample text using a TrueType font. Rendering to TIFF with compression can reduce file size.");

        // Define TIFF save options with compression to help reduce size.
        string tiffPath = "output.tiff";
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt4,
            Resolution = 300
        };

        // Render the document to a multipage TIFF file.
        doc.Save(tiffPath, saveOptions);

        // Verify that the TIFF file was created and report its size.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("TIFF file was not created.");

        long fileSize = new FileInfo(tiffPath).Length;
        Console.WriteLine($"TIFF file created at '{tiffPath}' with size {fileSize} bytes.");
    }
}

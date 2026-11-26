using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words TIFF with CCITT4 compression!");

        // Configure image save options for TIFF with CCITT4 compression.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.TiffCompression = TiffCompression.Ccitt4;

        // Save the document as a TIFF file.
        string outputPath = "output.tiff";
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the TIFF file.");

        Console.WriteLine("TIFF saved successfully: " + Path.GetFullPath(outputPath));
    }
}

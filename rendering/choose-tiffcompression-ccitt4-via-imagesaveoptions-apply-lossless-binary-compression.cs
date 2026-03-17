using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Define the folder where the output file will be saved.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);

        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add some content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for TIFF conversion.");
        // Insert an image (replace with a valid image path if needed).
        // builder.InsertImage("ImageDir/Logo.jpg");

        // Create ImageSaveOptions for TIFF format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);

        // Choose the CCITT4 compression scheme (lossless binary compression).
        options.TiffCompression = TiffCompression.Ccitt4;

        // Save the document as a TIFF image using the specified options.
        string outputPath = Path.Combine(artifactsDir, "Document_Ccitt4.tiff");
        doc.Save(outputPath, options);

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}

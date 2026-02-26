using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample content.
        builder.Writeln("This is a sample document rendered to TIFF.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page of the document.");

        // Create ImageSaveOptions for TIFF format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);

        // Set the TIFF compression type (LZW is the default, set explicitly here).
        options.TiffCompression = TiffCompression.Lzw;

        // Optional: set resolution and image size.
        options.Resolution = 300; // 300 DPI
        options.ImageSize = new Size(2480, 3508); // A4 at 300 DPI

        // Define output folder and file name.
        string artifactsDir = "Artifacts/";
        System.IO.Directory.CreateDirectory(artifactsDir);
        string outPath = System.IO.Path.Combine(artifactsDir, "DocumentRendered.tiff");

        // Save the document as a TIFF image using the specified options.
        doc.Save(outPath, options);
    }
}

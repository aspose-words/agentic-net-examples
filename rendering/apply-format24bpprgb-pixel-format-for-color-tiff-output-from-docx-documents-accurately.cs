using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple DOCX document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document rendered to a 24‑bpp color TIFF.");

        // Configure image save options for TIFF with 24‑bpp RGB pixel format.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            PixelFormat = ImagePixelFormat.Format24BppRgb
        };

        // Save the document as a TIFF file.
        string tiffPath = Path.Combine(outputDir, "Sample24bppRgb.tiff");
        doc.Save(tiffPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("TIFF file was not created.");

        // Optional: indicate success (no console input required).
        Console.WriteLine("TIFF saved successfully to: " + tiffPath);
    }
}

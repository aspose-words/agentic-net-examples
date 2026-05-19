using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple DOCX document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words!");
        builder.Writeln("This document will be saved as a color TIFF with 24‑bpp RGB pixel format.");

        // Configure TIFF save options with 24‑bpp RGB pixel format.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        tiffOptions.PixelFormat = ImagePixelFormat.Format24BppRgb;
        tiffOptions.Resolution = 300; // optional DPI setting

        // Save the document as a TIFF file.
        string tiffPath = Path.Combine(artifactsDir, "Document24bppRgb.tiff");
        doc.Save(tiffPath, tiffOptions);

        // Verify that the TIFF file was created and is not empty.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("TIFF file was not created.");

        if (new FileInfo(tiffPath).Length == 0)
            throw new InvalidOperationException("TIFF file is empty.");
    }
}

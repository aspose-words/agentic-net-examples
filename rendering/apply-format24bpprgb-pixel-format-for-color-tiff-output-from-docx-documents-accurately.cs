using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output folder and ensure it exists.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple DOCX document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words! This document will be rendered as a 24‑bpp color TIFF.");

        // Save the source DOCX (optional, just for reference).
        string docPath = Path.Combine(artifactsDir, "Sample.docx");
        doc.Save(docPath);

        // Configure image save options for TIFF with 24‑bpp RGB pixel format.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            PixelFormat = ImagePixelFormat.Format24BppRgb
        };

        // Render the document to a TIFF file.
        string tiffPath = Path.Combine(artifactsDir, "Sample_24bpp.tiff");
        doc.Save(tiffPath, tiffOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("TIFF file was not created.");

        Console.WriteLine($"TIFF saved successfully: {tiffPath}");
    }
}

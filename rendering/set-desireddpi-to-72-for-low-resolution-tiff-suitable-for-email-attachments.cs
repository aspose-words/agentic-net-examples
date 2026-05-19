using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple Word document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a low‑resolution TIFF example.");
        builder.Writeln("Suitable for email attachments.");

        // Prepare the folder for the output file.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string tiffPath = Path.Combine(outputDir, "LowResolution.tiff");

        // Configure image save options for TIFF with a low DPI (72).
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // The Resolution property sets both horizontal and vertical DPI.
            Resolution = 72f
        };

        // Save the document as a multipage TIFF.
        doc.Save(tiffPath, options);

        // Verify that the file was created.
        if (!File.Exists(tiffPath))
            throw new FileNotFoundException("The TIFF file was not created.", tiffPath);

        // Optionally, report the file size (useful for debugging).
        long fileSize = new FileInfo(tiffPath).Length;
        Console.WriteLine($"TIFF saved successfully to '{tiffPath}'. Size: {fileSize} bytes.");
    }
}

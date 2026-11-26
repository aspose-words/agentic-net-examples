using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the sample document and the output TIFF
        string docPath = "Sample.docx";
        string tiffPath = "LowResolution.tiff";

        // Create a simple Word document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document rendered to a low‑resolution TIFF (72 DPI).");

        // Save the DOCX (optional, just to demonstrate creation)
        doc.Save(docPath);

        // Desired DPI for the TIFF
        int desiredDpi = 72;

        // Configure TIFF save options with a low resolution (72 DPI)
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = desiredDpi
        };

        // Render the document to TIFF
        doc.Save(tiffPath, tiffOptions);

        // Verify that the TIFF file was created
        if (!File.Exists(tiffPath))
            throw new Exception("Failed to create the TIFF file.");

        // Output basic information about the generated file
        long fileSize = new FileInfo(tiffPath).Length;
        Console.WriteLine($"TIFF saved to '{tiffPath}' (Size: {fileSize} bytes, Resolution: {desiredDpi} DPI).");
    }
}

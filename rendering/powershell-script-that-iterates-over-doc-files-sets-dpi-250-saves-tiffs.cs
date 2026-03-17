using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Folder containing the source DOC/DOCX files
        string sourceFolder = @"C:\Docs\Input";
        // Folder where the generated TIFF files will be saved
        string outputFolder = @"C:\Docs\Output";

        // Ensure the output folder exists
        if (!Directory.Exists(outputFolder))
            Directory.CreateDirectory(outputFolder);

        // Get all .doc and .docx files (including subfolders)
        var files = Directory.GetFiles(sourceFolder, "*.*", SearchOption.AllDirectories);
        foreach (var filePath in files)
        {
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext != ".doc" && ext != ".docx")
                continue; // skip non‑Word files

            // Load the Word document
            Document doc = new Document(filePath);

            // Set up TIFF save options with 250 DPI resolution
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                Resolution = 250,
                // Optional: change compression if required
                // TiffCompression = TiffCompression.Lzw
            };

            // Build the output file name – same base name with .tiff extension
            string outputFileName = Path.ChangeExtension(Path.GetFileName(filePath), ".tiff");
            string outputPath = Path.Combine(outputFolder, outputFileName);

            // Save the document as a TIFF image
            doc.Save(outputPath, saveOptions);

            Console.WriteLine($"Converted '{filePath}' to TIFF at '{outputPath}' (250 DPI)");
        }
    }
}

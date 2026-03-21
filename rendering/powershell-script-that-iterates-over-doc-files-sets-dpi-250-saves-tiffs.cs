using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Use folders relative to the executable location
        string baseDir = AppContext.BaseDirectory;
        string sourceFolder = Path.Combine(baseDir, "Input");
        string outputFolder = Path.Combine(baseDir, "Output");

        // Ensure the folders exist
        Directory.CreateDirectory(sourceFolder);
        Directory.CreateDirectory(outputFolder);

        // Get all .doc and .docx files (including subfolders)
        var files = Directory.GetFiles(sourceFolder, "*.*", SearchOption.AllDirectories);
        if (files.Length == 0)
        {
            Console.WriteLine($"No Word documents found in '{sourceFolder}'. Place .doc or .docx files there and rerun.");
            return;
        }

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
                Resolution = 250
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

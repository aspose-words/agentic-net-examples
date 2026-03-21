using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class BatchDocxToTiff
{
    static void Main()
    {
        // Use folders relative to the executable location.
        string baseDir = AppContext.BaseDirectory;
        string sourceFolder = Path.Combine(baseDir, "Input");
        string outputFolder = Path.Combine(baseDir, "Output");

        // Ensure both directories exist.
        Directory.CreateDirectory(sourceFolder);
        Directory.CreateDirectory(outputFolder);

        // Shared ImageSaveOptions for all conversions.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Lzw,
            Resolution = 300
            // Multi-page layout setting omitted for compatibility.
        };

        string[] docxFiles = Directory.GetFiles(sourceFolder, "*.docx");
        if (docxFiles.Length == 0)
        {
            Console.WriteLine($"No DOCX files found in '{sourceFolder}'. Place files there and rerun.");
            return;
        }

        foreach (string docxPath in docxFiles)
        {
            Document doc = new Document(docxPath);
            string tiffPath = Path.Combine(outputFolder,
                Path.GetFileNameWithoutExtension(docxPath) + ".tiff");
            doc.Save(tiffPath, tiffOptions);
            Console.WriteLine($"Converted '{Path.GetFileName(docxPath)}' to '{Path.GetFileName(tiffPath)}'.");
        }
    }
}

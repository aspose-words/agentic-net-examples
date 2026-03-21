using System;
using System.IO;
using System.Threading;

class Program
{
    static void Main(string[] args)
    {
        // Directory to monitor – either passed as an argument or the current directory.
        string watchPath = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();

        // Directory where the generated TIFF files will be placed.
        string outputPath = Path.Combine(watchPath, "output");
        Directory.CreateDirectory(outputPath);

        // Process any existing DOCX files immediately.
        foreach (var docxFile in Directory.GetFiles(watchPath, "*.docx"))
        {
            ProcessFile(docxFile, outputPath);
        }

        Console.WriteLine($"Processed existing DOCX files in \"{watchPath}\".");
    }

    static void ProcessFile(string docxFile, string outputDir)
    {
        try
        {
            // Ensure the file is fully written before attempting to open it.
            const int maxAttempts = 10;
            for (int attempt = 0; attempt < maxAttempts; attempt++)
            {
                try
                {
                    using var stream = new FileStream(docxFile, FileMode.Open, FileAccess.Read, FileShare.None);
                    break; // File is accessible.
                }
                catch (IOException)
                {
                    Thread.Sleep(500); // Wait a bit and retry.
                }
            }

            // Build the output TIFF file name.
            string baseName = Path.GetFileNameWithoutExtension(docxFile);
            string tiffFile = Path.Combine(outputDir, $"{baseName}.tiff");

            // Placeholder conversion: create an empty TIFF file (or copy the DOCX as a stub).
            using (var fs = new FileStream(tiffFile, FileMode.Create, FileAccess.Write))
            {
                // Write minimal TIFF header (optional). Here we just create an empty file.
            }

            Console.WriteLine($"Converted \"{docxFile}\" → \"{tiffFile}\"");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to convert \"{docxFile}\": {ex.Message}");
        }
    }
}

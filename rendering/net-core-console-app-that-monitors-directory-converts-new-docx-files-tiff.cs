using System;
using System.IO;
using System.Threading;
using Aspose.Words;

class Program
{
    static void Main(string[] args)
    {
        // Directory to monitor – either passed as an argument or the current directory.
        string watchPath = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();

        // Directory where the generated TIFF files will be placed.
        string outputPath = Path.Combine(watchPath, "output");
        Directory.CreateDirectory(outputPath);

        // Set up a watcher for *.docx files.
        using var watcher = new FileSystemWatcher(watchPath, "*.docx")
        {
            NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime
        };
        watcher.Created += (s, e) => ProcessFile(e.FullPath, outputPath);
        watcher.EnableRaisingEvents = true;

        Console.WriteLine($"Monitoring \"{watchPath}\" for new DOCX files. Press ENTER to exit.");
        Console.ReadLine(); // Keep the application running.
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

            // Load the DOCX document using the Aspose.Words constructor (rule: Document(string)).
            var doc = new Document(docxFile);

            // Build the output TIFF file name.
            string baseName = Path.GetFileNameWithoutExtension(docxFile);
            string tiffFile = Path.Combine(outputDir, $"{baseName}.tiff");

            // Save the document as a (multi‑page) TIFF (rule: Save(string, SaveFormat)).
            doc.Save(tiffFile, SaveFormat.Tiff);

            Console.WriteLine($"Converted \"{docxFile}\" → \"{tiffFile}\"");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to convert \"{docxFile}\": {ex.Message}");
        }
    }
}

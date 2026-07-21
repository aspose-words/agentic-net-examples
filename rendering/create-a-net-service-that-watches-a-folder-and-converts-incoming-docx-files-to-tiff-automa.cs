using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class FolderWatcherService
{
    public static void Main()
    {
        // Set up input and output directories.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "WatchDemo");
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a sample DOCX file that the watcher will process.
        string sampleDocPath = Path.Combine(inputDir, "Sample.docx");
        CreateSampleDocument(sampleDocPath);

        // Convert the sample file immediately so the demo works even if the watcher
        // does not raise a Created event for a file that already exists.
        ConvertDocxToTiff(sampleDocPath, outputDir);

        // Watch the input folder for newly created DOCX files.
        using var watcher = new FileSystemWatcher(inputDir, "*.docx")
        {
            EnableRaisingEvents = true,
            IncludeSubdirectories = false
        };
        watcher.Created += (sender, e) => ConvertDocxToTiff(e.FullPath, outputDir);

        // Give the watcher time to detect and convert any new files.
        // In a real service this would run indefinitely; here we wait a short while.
        Task.Delay(5000).Wait();

        // Verify that the TIFF file was produced.
        string expectedTiff = Path.Combine(outputDir, "Sample.tiff");
        if (!File.Exists(expectedTiff))
            throw new Exception("TIFF conversion failed.");
    }

    // Generates a simple DOCX document for testing.
    private static void CreateSampleDocument(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document for automatic DOCX‑to‑TIFF conversion.");
        doc.Save(path);
    }

    // Converts a DOCX file to a single (multi‑page) TIFF image.
    private static void ConvertDocxToTiff(string docxPath, string outputDir)
    {
        // Retry opening the file in case it is still being written.
        for (int attempt = 0; attempt < 5; attempt++)
        {
            try
            {
                // Open the source document for reading.
                using var stream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                var document = new Document(stream);

                // Configure image save options for TIFF.
                var options = new ImageSaveOptions(SaveFormat.Tiff)
                {
                    Resolution = 300 // DPI – optional.
                };

                string tiffPath = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(docxPath) + ".tiff");
                document.Save(tiffPath, options);

                // Ensure the file was written.
                if (!File.Exists(tiffPath))
                    throw new Exception($"Failed to save TIFF for '{docxPath}'.");

                break; // Success.
            }
            catch (IOException)
            {
                // File may still be locked; wait briefly before retrying.
                Thread.Sleep(200);
            }
        }
    }
}

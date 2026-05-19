using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

public class FolderWatcherService
{
    public static void Main()
    {
        // Define input and output folders relative to the current directory.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Input");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a sample DOCX file that the watcher will process.
        string sampleDocPath = Path.Combine(inputFolder, "sample.docx");
        CreateSampleDocument(sampleDocPath);

        // Event used to signal that processing is complete.
        using var processingDone = new ManualResetEventSlim(false);

        // Set up a FileSystemWatcher to monitor the input folder for new DOCX files.
        using var watcher = new FileSystemWatcher(inputFolder, "*.docx")
        {
            EnableRaisingEvents = true,
            IncludeSubdirectories = false,
            NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime
        };

        // When a new DOCX file is created, convert it to a multipage TIFF.
        watcher.Created += (sender, args) => ConvertDocxToTiff(args.FullPath, outputFolder, processingDone);

        // Wait for the conversion to finish (or timeout after 10 seconds).
        processingDone.Wait(TimeSpan.FromSeconds(10));

        // Verify that the TIFF file was created.
        string expectedTiff = Path.Combine(outputFolder, "sample.tiff");
        if (File.Exists(expectedTiff))
        {
            Console.WriteLine($"TIFF successfully created at: {expectedTiff}");
        }
        else
        {
            Console.WriteLine("TIFF conversion failed or did not complete in time.");
        }
    }

    // Creates a simple multi‑page DOCX document for demonstration purposes.
    private static void CreateSampleDocument(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");

        doc.Save(filePath);
    }

    // Loads the DOCX file and saves it as a multipage TIFF image.
    private static void ConvertDocxToTiff(string docxPath, string outputFolder, ManualResetEventSlim doneEvent)
    {
        // The file may still be locked by the writer; retry until it becomes accessible.
        for (int attempt = 0; attempt < 10; attempt++)
        {
            try
            {
                using var stream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                var document = new Document(stream);

                var saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
                {
                    // Example: render at 300 DPI.
                    Resolution = 300
                };

                string tiffPath = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(docxPath) + ".tiff");
                document.Save(tiffPath, saveOptions);

                // Signal that processing is complete.
                doneEvent.Set();
                break;
            }
            catch (IOException)
            {
                // Wait briefly before retrying.
                Thread.Sleep(500);
            }
        }
    }
}

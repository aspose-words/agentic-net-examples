using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

public class FolderWatcherService
{
    private static readonly string InputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Input");
    private static readonly string OutputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
    private static readonly ManualResetEventSlim ProcessingCompleted = new ManualResetEventSlim(false);
    private static FileSystemWatcher _watcher;

    public static void Main()
    {
        // Ensure input and output directories exist.
        Directory.CreateDirectory(InputFolder);
        Directory.CreateDirectory(OutputFolder);

        // Create a sample DOCX file that the watcher will pick up.
        CreateSampleDocx(Path.Combine(InputFolder, "SampleDocument.docx"));

        // Set up the folder watcher.
        _watcher = new FileSystemWatcher(InputFolder, "*.docx")
        {
            EnableRaisingEvents = true,
            IncludeSubdirectories = false,
            NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime
        };
        _watcher.Created += OnDocxCreated;

        // Wait until the file has been processed or timeout after 15 seconds.
        if (!ProcessingCompleted.Wait(TimeSpan.FromSeconds(15)))
        {
            Console.WriteLine("Processing did not complete within the expected time.");
        }

        // Clean up.
        _watcher.Created -= OnDocxCreated;
        _watcher.Dispose();
    }

    private static void OnDocxCreated(object sender, FileSystemEventArgs e)
    {
        // The Created event may fire before the file is fully written; retry a few times.
        const int maxAttempts = 5;
        for (int attempt = 1; attempt <= maxAttempts; attempt++)
        {
            try
            {
                // Load the DOCX document.
                Document doc = new Document(e.FullPath);

                // Prepare TIFF save options.
                ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
                {
                    // Optional: set resolution or compression if desired.
                    Resolution = 300,
                    TiffCompression = TiffCompression.Lzw
                };

                // Build output file path.
                string outputFileName = Path.GetFileNameWithoutExtension(e.Name) + ".tiff";
                string outputPath = Path.Combine(OutputFolder, outputFileName);

                // Save the document as a multipage TIFF.
                doc.Save(outputPath, options);

                // Verify that the TIFF file was created.
                if (!File.Exists(outputPath))
                {
                    throw new InvalidOperationException($"Failed to create TIFF file at '{outputPath}'.");
                }

                Console.WriteLine($"Converted '{e.Name}' to TIFF successfully.");

                // Signal that processing is done.
                ProcessingCompleted.Set();
                break;
            }
            catch (IOException)
            {
                // File may still be locked; wait briefly before retrying.
                Thread.Sleep(500);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{e.Name}': {ex.Message}");
                ProcessingCompleted.Set();
                break;
            }
        }
    }

    private static void CreateSampleDocx(string filePath)
    {
        // Build a simple document with a few pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Page 1 - Hello World!");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 - Aspose.Words rendering example.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 - End of sample document.");

        // Save the sample DOCX.
        doc.Save(filePath);
    }
}

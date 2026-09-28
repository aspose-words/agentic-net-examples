using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Entry point
    public static void Main()
    {
        // Create a temporary folder to watch
        string watchFolder = Path.Combine(Path.GetTempPath(), "DocxToTiffWatcher_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(watchFolder);

        // Prepare a task that will be completed when a DOCX is converted
        var conversionCompleted = new TaskCompletionSource<bool>();

        // Set up the file system watcher
        using (var watcher = new FileSystemWatcher(watchFolder, "*.docx"))
        {
            watcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime;
            watcher.Created += (sender, e) => OnDocxCreated(e.FullPath, conversionCompleted);
            watcher.EnableRaisingEvents = true;

            // Create a sample DOCX file in the watched folder
            string sampleDocxPath = Path.Combine(watchFolder, "SampleDocument.docx");
            CreateSampleDocx(sampleDocxPath);

            // Wait for the conversion to finish (or timeout after 10 seconds)
            if (!conversionCompleted.Task.Wait(TimeSpan.FromSeconds(10)))
            {
                throw new Exception("DOCX to TIFF conversion did not complete in the expected time.");
            }
        }

        // Clean up temporary folder
        try { Directory.Delete(watchFolder, true); } catch { /* ignore cleanup errors */ }
    }

    // Handles the creation of a DOCX file
    private static void OnDocxCreated(string docxPath, TaskCompletionSource<bool> tcs)
    {
        // Process the file on a separate thread to avoid blocking the watcher
        ThreadPool.QueueUserWorkItem(_ =>
        {
            // Small delay to ensure the file is fully written
            Thread.Sleep(500);

            // Load the DOCX document
            var doc = new Document(docxPath);

            // Determine output TIFF path (same name, .tiff extension)
            string tiffPath = Path.ChangeExtension(docxPath, ".tiff");

            // Set up image save options for TIFF
            var saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Ensure all pages are saved; default behavior already does this
                // No additional configuration required
            };

            // Save the document as TIFF
            doc.Save(tiffPath, saveOptions);

            // Validate that the TIFF file was created and has content
            if (!File.Exists(tiffPath) || new FileInfo(tiffPath).Length == 0)
            {
                tcs.TrySetException(new Exception("Failed to create TIFF output."));
                return;
            }

            // Signal successful conversion
            tcs.TrySetResult(true);
        });
    }

    // Creates a simple DOCX file with sample content
    private static void CreateSampleDocx(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document generated for automatic DOCX to TIFF conversion.");
        doc.Save(path);
    }
}

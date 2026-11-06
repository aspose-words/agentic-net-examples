using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static async Task Main()
    {
        // Create a temporary folder to watch.
        string watchFolder = Path.Combine(Path.GetTempPath(), "AsposeWatch_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(watchFolder);

        // TaskCompletionSource to signal when conversion finishes.
        var conversionDone = new TaskCompletionSource<bool>();

        // Set up a FileSystemWatcher for *.docx files.
        using var watcher = new FileSystemWatcher(watchFolder, "*.docx")
        {
            EnableRaisingEvents = true,
            IncludeSubdirectories = false,
            NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime
        };

        watcher.Created += (sender, e) =>
        {
            try
            {
                // Load the newly created DOCX file.
                var doc = new Document(e.FullPath);

                // Define the output TIFF path.
                string tiffPath = Path.ChangeExtension(e.FullPath, ".tiff");

                // Save the document as a multipage TIFF.
                var saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
                doc.Save(tiffPath, saveOptions);

                // Verify that the TIFF file was created and is not empty.
                if (!File.Exists(tiffPath) || new FileInfo(tiffPath).Length == 0)
                    throw new InvalidOperationException("TIFF conversion failed or produced an empty file.");

                // Signal successful conversion.
                conversionDone.TrySetResult(true);
            }
            catch (Exception ex)
            {
                // Propagate any errors to the awaiting task.
                conversionDone.TrySetException(ex);
            }
        };

        // Ensure the watcher is ready before creating the sample file.
        await Task.Delay(500);

        // Create a sample DOCX document in the watched folder.
        string sampleDocxPath = Path.Combine(watchFolder, "Sample.docx");
        var sampleDoc = new Document();
        var builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Hello, Aspose.Words! This document will be converted to TIFF automatically.");
        sampleDoc.Save(sampleDocxPath, SaveFormat.Docx);

        // Wait for the conversion to complete or timeout after a reasonable period.
        var finishedTask = await Task.WhenAny(conversionDone.Task, Task.Delay(5000));
        if (finishedTask != conversionDone.Task)
            throw new TimeoutException("TIFF conversion did not complete within the expected time.");

        // Optional cleanup: delete the temporary folder and its contents.
        // Directory.Delete(watchFolder, true);
    }
}

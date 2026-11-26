using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class DocxToTiffWatcher
{
    // Entry point
    public static void Main()
    {
        // Create a temporary folder to watch
        string watchFolder = Path.Combine(Path.GetTempPath(),
            "DocxWatchFolder_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(watchFolder);

        // Set up a FileSystemWatcher for *.docx files
        using (FileSystemWatcher watcher = new FileSystemWatcher())
        {
            watcher.Path = watchFolder;
            watcher.Filter = "*.docx";
            watcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime;
            watcher.Created += OnDocxCreated;
            watcher.EnableRaisingEvents = true;

            // Create a sample DOCX file in the watched folder
            string sampleDocxPath = Path.Combine(watchFolder, "SampleDocument.docx");
            CreateSampleDocx(sampleDocxPath);

            // Allow some time for the watcher to process the file
            Task.Delay(3000).Wait();

            // Disable events and clean up
            watcher.EnableRaisingEvents = false;
        }

        // Verify that the TIFF was created
        string expectedTiffPath = Path.Combine(watchFolder, "SampleDocument.tiff");
        if (!File.Exists(expectedTiffPath))
            throw new InvalidOperationException($"TIFF file was not created: {expectedTiffPath}");

        // Optional: clean up temporary folder
        // Directory.Delete(watchFolder, true);
    }

    // Event handler that converts a newly created DOCX to TIFF
    private static void OnDocxCreated(object sender, FileSystemEventArgs e)
    {
        // Wait briefly to give the OS a chance to finish writing the file
        Task.Delay(500).Wait();

        // Try to load the document, retrying if the file is still locked
        Document doc = null;
        const int maxAttempts = 5;
        int attempt = 0;
        while (attempt < maxAttempts)
        {
            try
            {
                doc = new Document(e.FullPath);
                break; // Success
            }
            catch (IOException)
            {
                // File is still in use; wait and retry
                attempt++;
                Task.Delay(500).Wait();
            }
        }

        if (doc == null)
            throw new InvalidOperationException($"Unable to load DOCX file after multiple attempts: {e.FullPath}");

        // Determine output TIFF path (same name, .tiff extension)
        string tiffPath = Path.ChangeExtension(e.FullPath, ".tiff");

        // Set up image save options for TIFF (multipage TIFF by default)
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);

        // Save the document as TIFF
        doc.Save(tiffPath, saveOptions);

        // Verify that the TIFF file exists
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
    }

    // Helper method to create a simple DOCX document
    private static void CreateSampleDocx(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document generated for the watcher test.");
        doc.Save(path);
    }
}

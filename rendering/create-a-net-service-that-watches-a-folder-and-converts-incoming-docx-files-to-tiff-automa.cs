using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Event used to signal that a file has been processed.
    private static readonly ManualResetEventSlim _processedEvent = new ManualResetEventSlim(false);

    public static void Main()
    {
        // Define folders for input DOCX files and output TIFF files.
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "Input");
        string outputFolder = Path.Combine(baseDir, "Output");

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Set up a watcher that reacts to newly created DOCX files.
        using (FileSystemWatcher watcher = new FileSystemWatcher(inputFolder, "*.docx"))
        {
            watcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime;
            watcher.Created += (sender, e) => OnCreated(e.FullPath, outputFolder);
            watcher.EnableRaisingEvents = true;

            // Create a sample DOCX file to trigger the watcher.
            string sampleDocPath = Path.Combine(inputFolder, "SampleDocument.docx");
            CreateSampleDocument(sampleDocPath);

            // Wait until the file has been processed or timeout after 10 seconds.
            if (!_processedEvent.Wait(TimeSpan.FromSeconds(10)))
                throw new InvalidOperationException("The DOCX file was not processed in time.");

            // Verify that the TIFF output exists.
            string expectedTiffPath = Path.Combine(outputFolder, "SampleDocument.tiff");
            if (!File.Exists(expectedTiffPath))
                throw new FileNotFoundException("Expected TIFF output was not created.", expectedTiffPath);
        }
    }

    // Creates a minimal DOCX document with some text.
    private static void CreateSampleDocument(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document generated for the folder‑watcher example.");
        doc.Save(path);
    }

    // Waits until the file can be opened for reading (i.e., the writer has released the handle).
    private static void WaitForFileReady(string filePath, int timeoutMs = 5000)
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();
        while (true)
        {
            try
            {
                using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    // If we can open the file, it is ready.
                    break;
                }
            }
            catch (IOException)
            {
                if (sw.ElapsedMilliseconds > timeoutMs)
                    throw new TimeoutException($"Timed out waiting for file '{filePath}' to become ready.");
                Thread.Sleep(100); // Small pause before retry.
            }
        }
    }

    // Handles the creation of a new DOCX file: converts it to a multi‑page TIFF.
    private static void OnCreated(string docxPath, string outputFolder)
    {
        try
        {
            // Ensure the file is fully written and not locked.
            WaitForFileReady(docxPath);

            // Load the newly created document.
            Document doc = new Document(docxPath);

            // Configure image save options for TIFF output.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Render all pages into a single multi‑frame TIFF.
                PageLayout = MultiPageLayout.TiffFrames()
            };

            // Determine output file name (same base name, .tiff extension).
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(docxPath);
            string tiffPath = Path.Combine(outputFolder, fileNameWithoutExt + ".tiff");

            // Save the document as TIFF.
            doc.Save(tiffPath, options);
        }
        finally
        {
            // Signal that processing is complete so the program can exit.
            _processedEvent.Set();
        }
    }
}

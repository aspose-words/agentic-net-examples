using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Event used to signal that a file has been processed.
    private static readonly ManualResetEventSlim ProcessedEvent = new ManualResetEventSlim(false);

    public static void Main()
    {
        // Define input and output folders relative to the current directory.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Input");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Set up a watcher that looks for newly created DOCX files.
        using (FileSystemWatcher watcher = new FileSystemWatcher(inputFolder, "*.docx"))
        {
            watcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite;
            watcher.Created += (sender, e) => OnNewDocx(e.FullPath, outputFolder);
            watcher.EnableRaisingEvents = true;

            // Create a sample DOCX file to trigger the watcher.
            CreateSampleDocx(Path.Combine(inputFolder, "Sample.docx"));

            // Wait until the file is processed or timeout after 10 seconds.
            ProcessedEvent.Wait(TimeSpan.FromSeconds(10));
        }
    }

    // Called when a new DOCX file appears in the monitored folder.
    private static void OnNewDocx(string docxPath, string outputFolder)
    {
        try
        {
            // Load the newly created document.
            Document doc = new Document(docxPath);

            // Configure image save options for TIFF output.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Optional: set resolution (dpi) if desired.
                Resolution = 300
            };

            // Build the output file path with the same name but .tiff extension.
            string tiffPath = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(docxPath) + ".tiff");

            // Save the document as a multipage TIFF.
            doc.Save(tiffPath, options);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");

            // Signal that processing is complete.
            ProcessedEvent.Set();
        }
        catch (Exception ex)
        {
            // In a real application you would handle logging here.
            Console.Error.WriteLine($"Error processing '{docxPath}': {ex.Message}");
            ProcessedEvent.Set(); // Ensure the program can exit even on error.
        }
    }

    // Generates a simple DOCX file with a couple of pages.
    private static void CreateSampleDocx(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("This is the first page.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is the second page.");

        doc.Save(filePath);
    }
}

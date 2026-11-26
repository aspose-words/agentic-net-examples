using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Entry point of the console application.
    public static void Main()
    {
        // Define input and output directories relative to the current working directory.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");

        // Ensure the directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Set up a synchronization primitive to wait for processing or timeout.
        using (var processedEvent = new ManualResetEventSlim(false))
        {
            // Configure a FileSystemWatcher to monitor the input directory for new DOCX files.
            using (var watcher = new FileSystemWatcher(inputDir, "*.docx"))
            {
                watcher.Created += (sender, e) =>
                {
                    try
                    {
                        // Load the newly created DOCX document.
                        var doc = new Document(e.FullPath);

                        // Prepare TIFF rendering options (default renders all pages).
                        var tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);

                        // Determine the output TIFF file path.
                        string outputTiffPath = Path.Combine(outputDir,
                            Path.GetFileNameWithoutExtension(e.Name) + ".tiff");

                        // Save the document as TIFF.
                        doc.Save(outputTiffPath, tiffOptions);

                        // Verify that the TIFF file was created.
                        if (!File.Exists(outputTiffPath))
                            throw new InvalidOperationException($"Failed to create TIFF file: {outputTiffPath}");

                        // Signal that processing is complete.
                        processedEvent.Set();
                    }
                    catch (Exception ex)
                    {
                        // Propagate any exception to the main thread.
                        Console.Error.WriteLine($"Error processing file '{e.Name}': {ex.Message}");
                        processedEvent.Set();
                    }
                };

                // Begin watching.
                watcher.EnableRaisingEvents = true;

                // Create a sample DOCX file in the input directory after the watcher is active.
                string sampleDocxPath = Path.Combine(inputDir, "SampleDocument.docx");
                CreateSampleDocx(sampleDocxPath);

                // Wait for the processing to complete or timeout after 10 seconds.
                bool signaled = processedEvent.Wait(TimeSpan.FromSeconds(10));

                // If the event was not signaled, it means the file was not processed in time.
                if (!signaled)
                    throw new TimeoutException("Timed out waiting for DOCX to TIFF conversion.");
            }
        }

        // At this point, the conversion should have succeeded.
        // The application exits automatically.
    }

    // Helper method to create a simple DOCX document with sample content.
    private static void CreateSampleDocx(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document generated for TIFF conversion.");
        builder.Writeln("It contains multiple lines to demonstrate rendering.");
        doc.Save(path);
    }
}

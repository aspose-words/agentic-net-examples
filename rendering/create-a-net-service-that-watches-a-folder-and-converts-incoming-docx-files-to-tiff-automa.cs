using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Input and output folders relative to the current working directory.
    private static readonly string BaseDir = Path.Combine(Directory.GetCurrentDirectory(), "WatchFolder");
    private static readonly string InputDir = Path.Combine(BaseDir, "Input");
    private static readonly string OutputDir = Path.Combine(BaseDir, "Output");

    public static void Main()
    {
        // Ensure the folders exist.
        Directory.CreateDirectory(InputDir);
        Directory.CreateDirectory(OutputDir);

        // Set up a watcher for newly created DOCX files.
        using (var watcher = new FileSystemWatcher(InputDir, "*.docx"))
        {
            watcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime;
            watcher.Created += OnCreated;
            watcher.EnableRaisingEvents = true;

            // Create a sample DOCX file to trigger the conversion.
            string sampleDocPath = Path.Combine(InputDir, "SampleDocument.docx");
            CreateSampleDocument(sampleDocPath);

            // Give the watcher time to process the file.
            Thread.Sleep(3000);
        }
    }

    // Event handler that converts the detected DOCX file to a multipage TIFF.
    private static void OnCreated(object sender, FileSystemEventArgs e)
    {
        // Wait until the file is fully written and can be opened.
        const int maxAttempts = 10;
        int attempts = 0;
        while (attempts < maxAttempts)
        {
            try
            {
                using (FileStream testStream = File.Open(e.FullPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    // If we can open it, break out of the loop.
                    break;
                }
            }
            catch (IOException)
            {
                attempts++;
                Thread.Sleep(200);
            }
        }

        // Load the DOCX document.
        Document doc = new Document(e.FullPath);

        // Prepare the output TIFF path.
        string outputFileName = Path.GetFileNameWithoutExtension(e.Name) + ".tiff";
        string outputPath = Path.Combine(OutputDir, outputFileName);

        // Configure image save options for TIFF.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render all pages into a single multipage TIFF.
            PageSet = new PageSet(0), // Start from the first page.
            Resolution = 300 // 300 DPI for decent quality.
        };

        // Save the document as TIFF.
        doc.Save(outputPath, options);

        // Validate that the TIFF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create TIFF file at '{outputPath}'.");
    }

    // Creates a simple DOCX file with some text content.
    private static void CreateSampleDocument(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document generated for automatic conversion.");
        builder.Writeln("It will be converted to a TIFF image by the watcher service.");
        doc.Save(path);
    }
}

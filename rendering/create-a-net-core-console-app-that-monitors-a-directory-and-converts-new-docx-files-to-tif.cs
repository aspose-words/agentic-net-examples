using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main(string[] args)
    {
        // Define the folder to watch.
        string watchFolder = Path.Combine(Directory.GetCurrentDirectory(), "WatchFolder");
        Directory.CreateDirectory(watchFolder);

        // Create a sample DOCX file if none exist.
        string sampleDocPath = Path.Combine(watchFolder, "Sample.docx");
        if (!File.Exists(sampleDocPath))
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("This is a sample document created for TIFF conversion.");
            sampleDoc.Save(sampleDocPath);
        }

        // Process any existing DOCX files.
        foreach (string docxPath in Directory.GetFiles(watchFolder, "*.docx"))
        {
            ConvertDocxToTiff(docxPath);
        }

        // Set up a watcher to handle newly created DOCX files.
        using (FileSystemWatcher watcher = new FileSystemWatcher(watchFolder, "*.docx"))
        {
            watcher.Created += (sender, e) =>
            {
                // Small delay to ensure the file is fully written.
                Thread.Sleep(500);
                ConvertDocxToTiff(e.FullPath);
            };
            watcher.EnableRaisingEvents = true;

            // Monitor for a short, deterministic period then exit.
            Thread.Sleep(5000);
        }
    }

    private static void ConvertDocxToTiff(string docxPath)
    {
        // Ensure the source file exists.
        if (!File.Exists(docxPath))
            throw new FileNotFoundException($"Source DOCX not found: {docxPath}");

        // Load the document.
        Document doc = new Document(docxPath);

        // Prepare TIFF save options.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        // Render all pages; no additional configuration needed.

        // Determine output path.
        string tiffPath = Path.ChangeExtension(docxPath, ".tiff");

        // Save as TIFF.
        doc.Save(tiffPath, options);

        // Validate that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
    }
}

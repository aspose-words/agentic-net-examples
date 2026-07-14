using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Directory to watch for new DOCX files.
        string watchDir = Path.Combine(Directory.GetCurrentDirectory(), "WatchFolder");
        Directory.CreateDirectory(watchDir);

        // Create a sample DOCX file so the watcher has something to process.
        string sampleDocPath = Path.Combine(watchDir, "Sample.docx");
        CreateSampleDocument(sampleDocPath);

        // Set up a FileSystemWatcher to react to newly created DOCX files.
        using var watcher = new FileSystemWatcher(watchDir, "*.docx")
        {
            EnableRaisingEvents = true,
            IncludeSubdirectories = false
        };
        watcher.Created += (sender, args) => ConvertDocxToTiff(args.FullPath);

        // Process any DOCX files that already exist in the folder.
        foreach (string existingFile in Directory.GetFiles(watchDir, "*.docx"))
        {
            ConvertDocxToTiff(existingFile);
        }

        // Allow a short period for the watcher to catch any additional files,
        // then exit automatically.
        Thread.Sleep(2000);
    }

    // Generates a simple DOCX document with some text.
    private static void CreateSampleDocument(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words!");
        doc.Save(filePath);
    }

    // Converts the specified DOCX file to a TIFF image using Aspose.Words.
    private static void ConvertDocxToTiff(string docxPath)
    {
        try
        {
            // Load the DOCX document.
            var doc = new Document(docxPath);

            // Configure image save options for TIFF output.
            var options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Example: set resolution to 300 DPI.
                Resolution = 300
            };

            // Determine the output TIFF file path.
            string tiffPath = Path.ChangeExtension(docxPath, ".tiff");

            // Save the document as a TIFF image.
            doc.Save(tiffPath, options);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"TIFF file was not created: {tiffPath}");
        }
        catch (Exception ex)
        {
            // Log any errors to the standard error stream.
            Console.Error.WriteLine($"Error converting '{docxPath}' to TIFF: {ex.Message}");
        }
    }
}

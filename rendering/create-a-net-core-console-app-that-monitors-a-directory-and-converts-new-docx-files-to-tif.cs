using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define input and output folders relative to the executable location.
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Set up a watcher that reacts to newly created DOCX files.
        using (FileSystemWatcher watcher = new FileSystemWatcher())
        {
            watcher.Path = inputDir;
            watcher.Filter = "*.docx";
            watcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime;
            watcher.Created += (sender, e) => ConvertDocxToTiff(e.FullPath, outputDir);
            watcher.EnableRaisingEvents = true;

            // Create a sample DOCX file that will trigger the conversion.
            string sampleDocPath = Path.Combine(inputDir, "SampleDocument.docx");
            CreateSampleDocument(sampleDocPath);

            // Give the watcher time to detect and process the file.
            Thread.Sleep(3000);
        }

        // Verify that the TIFF file was produced.
        string expectedTiff = Path.Combine(outputDir, "SampleDocument.tiff");
        if (!File.Exists(expectedTiff))
            throw new InvalidOperationException("TIFF conversion failed: output file not found.");
    }

    // Generates a simple two‑page DOCX document.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("First page of the sample document.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page of the sample document.");

        doc.Save(filePath, SaveFormat.Docx);
    }

    // Converts a DOCX file to a multi‑page TIFF image.
    private static void ConvertDocxToTiff(string docxPath, string outputDir)
    {
        const int maxAttempts = 5;
        int attempt = 0;

        while (true)
        {
            try
            {
                // Load the source document.
                Document doc = new Document(docxPath);

                // Configure image save options for TIFF output.
                ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
                {
                    // Render all pages into a single multi‑page TIFF.
                    PageSet = new PageSet(0, doc.PageCount - 1),
                    Resolution = 300 // 300 DPI for good quality.
                };

                // Build the output file path.
                string tiffFileName = Path.GetFileNameWithoutExtension(docxPath) + ".tiff";
                string tiffPath = Path.Combine(outputDir, tiffFileName);

                // Save the document as TIFF.
                doc.Save(tiffPath, options);
                break; // Success – exit the retry loop.
            }
            catch (IOException) when (attempt < maxAttempts)
            {
                // The file may still be locked by the OS; wait briefly and retry.
                attempt++;
                Thread.Sleep(200);
            }
            catch (Exception ex)
            {
                // Log any other errors.
                Console.Error.WriteLine($"Error converting '{docxPath}': {ex.Message}");
                break;
            }
        }
    }
}

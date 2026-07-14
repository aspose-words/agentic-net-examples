using System;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input DOC files and output TIFF files.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputTiffs");

        // Ensure clean directories.
        if (Directory.Exists(inputDir)) Directory.Delete(inputDir, true);
        if (Directory.Exists(outputDir)) Directory.Delete(outputDir, true);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a set of sample DOC files.
        int docCount = 5; // Number of documents to generate.
        List<string> inputFiles = new List<string>();
        for (int i = 1; i <= docCount; i++)
        {
            string docPath = Path.Combine(inputDir, $"Sample{i}.doc");
            CreateSampleDocument(docPath, i);
            inputFiles.Add(docPath);
        }

        // Parallel conversion of DOC files to multi‑page TIFF.
        Parallel.ForEach(inputFiles, inputPath =>
        {
            // Load the source document.
            Document doc = new Document(inputPath);

            // Configure image save options for TIFF.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Render all pages into a single multi‑frame TIFF.
                PageLayout = MultiPageLayout.TiffFrames(),
                // Example resolution; adjust as needed.
                Resolution = 300
            };

            // Determine output file name.
            string fileName = Path.GetFileNameWithoutExtension(inputPath);
            string tiffPath = Path.Combine(outputDir, $"{fileName}.tiff");

            // Save the document as TIFF.
            doc.Save(tiffPath, options);
        });

        // Verify that all TIFF files were created.
        foreach (string inputPath in inputFiles)
        {
            string fileName = Path.GetFileNameWithoutExtension(inputPath);
            string tiffPath = Path.Combine(outputDir, $"{fileName}.tiff");
            if (!File.Exists(tiffPath))
                throw new FileNotFoundException($"Failed to create TIFF for {inputPath}");
        }

        // Indicate successful completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }

    // Helper method to create a simple DOC file with multiple pages.
    private static void CreateSampleDocument(string filePath, int documentNumber)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add several pages with identifiable content.
        for (int page = 1; page <= 3; page++)
        {
            builder.Writeln($"Document {documentNumber} - Page {page}");
            if (page < 3)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document.
        doc.Save(filePath);
    }
}

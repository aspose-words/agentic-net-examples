using System;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class ParallelDocToTiffConverter
{
    public static void Main()
    {
        // Define folders for input DOC files and output TIFF files.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "ConversionDemo");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputTiffs");

        // Ensure clean directories.
        if (Directory.Exists(baseDir))
            Directory.Delete(baseDir, true);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a set of sample DOC files.
        int docCount = 5;
        List<string> sourceFiles = new List<string>();
        for (int i = 1; i <= docCount; i++)
        {
            string docPath = Path.Combine(inputDir, $"Sample{i}.docx");
            CreateSampleDocument(docPath, i);
            sourceFiles.Add(docPath);
        }

        // Parallel conversion of each DOC file to a multipage TIFF.
        Parallel.ForEach(sourceFiles, sourcePath =>
        {
            // Load the document.
            Document doc = new Document(sourcePath);

            // Configure image save options for TIFF.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Render all pages into a single multi‑frame TIFF.
                PageLayout = MultiPageLayout.TiffFrames(),
                // Optional: set resolution for better quality.
                Resolution = 300
            };

            // Determine output file name.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
            string tiffPath = Path.Combine(outputDir, $"{fileNameWithoutExt}.tiff");

            // Save the document as TIFF.
            doc.Save(tiffPath, options);
        });

        // Verify that all TIFF files were created.
        foreach (string sourcePath in sourceFiles)
        {
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
            string tiffPath = Path.Combine(outputDir, $"{fileNameWithoutExt}.tiff");
            if (!File.Exists(tiffPath))
                throw new FileNotFoundException($"Failed to create TIFF for '{sourcePath}'.");
        }

        // Indicate successful completion (no interactive output required).
        Console.WriteLine("All documents have been converted to TIFF successfully.");
    }

    // Helper method to create a simple DOCX file with sample content.
    private static void CreateSampleDocument(string filePath, int index)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln($"This is sample document #{index}.");
        builder.Writeln("It contains multiple pages for conversion testing.");
        // Insert page breaks to generate several pages.
        for (int p = 1; p <= 3; p++)
        {
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Page {p + 1} of document #{index}.");
        }

        doc.Save(filePath);
    }
}

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
        // Base temporary directory for the demo
        string baseDir = Path.Combine(Path.GetTempPath(), "AsposeWordsBatchDemo");
        string sourceDir = Path.Combine(baseDir, "SourceDocs");
        string outputDir = Path.Combine(baseDir, "TiffOutputs");

        // Ensure clean directories
        if (Directory.Exists(baseDir))
            Directory.Delete(baseDir, true);
        Directory.CreateDirectory(sourceDir);
        Directory.CreateDirectory(outputDir);

        // Number of sample documents to create
        int docCount = 5;

        // Create sample DOCX files
        List<string> sourceFiles = new List<string>();
        for (int i = 1; i <= docCount; i++)
        {
            string filePath = Path.Combine(sourceDir, $"SampleDocument{i}.docx");
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add enough content to generate multiple pages
            builder.Writeln($"--- Sample Document {i} ---");
            for (int line = 1; line <= 200; line++)
            {
                builder.Writeln($"Line {line} of document {i}");
            }

            doc.Save(filePath);
            sourceFiles.Add(filePath);
        }

        // Parallel conversion of DOCX files to multipage TIFF
        Parallel.ForEach(sourceFiles, sourcePath =>
        {
            // Load the source document
            Document doc = new Document(sourcePath);

            // Prepare TIFF save options
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Ensure all pages are saved into a single multipage TIFF
                PageSet = new PageSet(0, doc.PageCount - 1)
            };

            // Determine output file path
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
            string tiffPath = Path.Combine(outputDir, $"{fileNameWithoutExt}.tiff");

            // Save the document as TIFF
            doc.Save(tiffPath, saveOptions);

            // Validate that the TIFF file was created
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF for '{sourcePath}'.");
        });

        // Final verification: ensure all TIFF files exist
        foreach (var src in sourceFiles)
        {
            string tiffPath = Path.Combine(outputDir, $"{Path.GetFileNameWithoutExtension(src)}.tiff");
            if (!File.Exists(tiffPath))
                throw new FileNotFoundException($"Expected TIFF file not found: {tiffPath}");
        }

        // Optionally, output the result count (no user interaction required)
        Console.WriteLine($"Converted {docCount} documents to TIFF. Output folder: {outputDir}");
    }
}

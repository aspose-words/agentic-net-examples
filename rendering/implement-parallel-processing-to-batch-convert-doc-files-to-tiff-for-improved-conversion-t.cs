using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class ParallelDocToTiffConverter
{
    public static void Main()
    {
        // Define folders for source DOC files and resulting TIFF files.
        string baseDir = Directory.GetCurrentDirectory();
        string sourceDir = Path.Combine(baseDir, "SourceDocs");
        string outputDir = Path.Combine(baseDir, "OutputTiffs");

        // Ensure the directories exist.
        Directory.CreateDirectory(sourceDir);
        Directory.CreateDirectory(outputDir);

        // Create a set of sample DOC files locally.
        const int sampleCount = 5;
        for (int i = 1; i <= sampleCount; i++)
        {
            string docPath = Path.Combine(sourceDir, $"SampleDocument{i}.doc");
            CreateSampleDocument(docPath, i);
        }

        // Get all DOC files that need to be converted.
        string[] docFiles = Directory.GetFiles(sourceDir, "*.doc");

        // Convert each DOC to a multi‑page TIFF in parallel.
        Parallel.ForEach(docFiles, docFile =>
        {
            // Load the source document.
            Document doc = new Document(docFile);

            // Configure image save options for TIFF output.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Render all pages into separate frames of a single TIFF file.
                PageLayout = MultiPageLayout.TiffFrames(),
                // Optional: set a higher resolution for better quality.
                Resolution = 300
            };

            // Determine the output TIFF file path.
            string tiffFileName = Path.GetFileNameWithoutExtension(docFile) + ".tiff";
            string tiffPath = Path.Combine(outputDir, tiffFileName);

            // Save the document as TIFF.
            doc.Save(tiffPath, saveOptions);
        });

        // Verify that each TIFF file was created.
        int successCount = 0;
        foreach (string docFile in docFiles)
        {
            string expectedTiff = Path.Combine(outputDir,
                Path.GetFileNameWithoutExtension(docFile) + ".tiff");
            if (File.Exists(expectedTiff))
                successCount++;
        }

        Console.WriteLine($"Converted {successCount} out of {docFiles.Length} documents to TIFF.");
    }

    // Helper method to create a simple multi‑page DOC file.
    private static void CreateSampleDocument(string filePath, int index)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some content and page breaks to generate multiple pages.
        builder.Writeln($"Sample Document #{index}");
        builder.Writeln("This is the first page of the document.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is the second page of the document.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is the third page of the document.");

        // Save the document in the legacy DOC format.
        doc.Save(filePath, SaveFormat.Doc);
    }
}

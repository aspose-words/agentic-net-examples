using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class ParallelDocToTiffConverter
{
    public static void Main()
    {
        // Create a temporary working directory.
        string workDir = Path.Combine(Path.GetTempPath(), "AsposeWordsParallelDemo");
        Directory.CreateDirectory(workDir);

        // Create a subfolder for source DOC files.
        string sourceDir = Path.Combine(workDir, "SourceDocs");
        Directory.CreateDirectory(sourceDir);

        // Create a subfolder for the resulting TIFF files.
        string outputDir = Path.Combine(workDir, "TiffOutputs");
        Directory.CreateDirectory(outputDir);

        // Generate a few sample DOC files.
        const int sampleCount = 5;
        for (int i = 1; i <= sampleCount; i++)
        {
            string docPath = Path.Combine(sourceDir, $"SampleDocument{i}.doc");
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Sample Document {i} - Page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Sample Document {i} - Page 2.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Sample Document {i} - Page 3.");

            doc.Save(docPath);
        }

        // Gather all DOC files to be processed.
        string[] docFiles = Directory.GetFiles(sourceDir, "*.doc");

        // Convert each DOC to a multipage TIFF in parallel.
        Parallel.ForEach(docFiles, docFile =>
        {
            // Load the source document.
            Document document = new Document(docFile);

            // Prepare the output TIFF path.
            string tiffFileName = Path.GetFileNameWithoutExtension(docFile) + ".tiff";
            string tiffPath = Path.Combine(outputDir, tiffFileName);

            // Configure image save options for a multipage TIFF.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Render all pages into a single multi‑frame TIFF.
                PageLayout = MultiPageLayout.TiffFrames(),
                // Set a reasonable resolution (dots per inch).
                Resolution = 300
            };

            // Save the document as TIFF.
            document.Save(tiffPath, saveOptions);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
        });

        // Simple verification: list the generated TIFF files.
        Console.WriteLine("Conversion completed. Generated TIFF files:");
        foreach (string tiffFile in Directory.GetFiles(outputDir, "*.tiff"))
        {
            Console.WriteLine(tiffFile);
        }

        // Cleanup (optional): delete the temporary working directory.
        // Directory.Delete(workDir, true);
    }
}

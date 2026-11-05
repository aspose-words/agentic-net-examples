using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Working folder for sample DOC files and output TIFF files.
        string workFolder = Path.Combine(Path.GetTempPath(), "AsposeBatchConvert");
        Directory.CreateDirectory(workFolder);

        // Clean any previous files.
        foreach (string file in Directory.GetFiles(workFolder))
            File.Delete(file);

        // Number of sample DOC files to create.
        const int sampleCount = 5;

        // Create sample DOC files with simple content.
        for (int i = 1; i <= sampleCount; i++)
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add paragraphs and page breaks.
            DocumentBuilder builder = new DocumentBuilder(doc);
            for (int p = 1; p <= 3; p++)
            {
                builder.Writeln($"Document {i}, page {p} content.");
                if (p < 3)
                {
                    // Insert a page break between pages.
                    builder.InsertBreak(BreakType.PageBreak);
                }
            }

            // Save the DOC file.
            string docPath = Path.Combine(workFolder, $"Sample{i}.doc");
            doc.Save(docPath, SaveFormat.Doc);
        }

        // Get all DOC files in the working folder.
        string[] docFiles = Directory.GetFiles(workFolder, "*.doc");

        // Parallel conversion of each DOC file to a multipage TIFF.
        ParallelOptions parallelOptions = new ParallelOptions
        {
            MaxDegreeOfParallelism = Environment.ProcessorCount
        };

        Parallel.ForEach(docFiles, parallelOptions, docPath =>
        {
            // Load the source document.
            Document srcDoc = new Document(docPath);

            // Prepare TIFF output path.
            string tiffPath = Path.ChangeExtension(docPath, ".tiff");

            // Configure image save options for TIFF.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Render all pages into a single multipage TIFF.
                PageSet = PageSet.All
            };

            // Save the document as TIFF.
            srcDoc.Save(tiffPath, saveOptions);

            // Validate that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
        });

        // Verify that the expected number of TIFF files were produced.
        int tiffCount = Directory.GetFiles(workFolder, "*.tiff").Length;
        if (tiffCount != sampleCount)
            throw new InvalidOperationException($"Expected {sampleCount} TIFF files, but found {tiffCount}.");
    }
}

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

class BatchDocToTiffConverter
{
    // Converts all .doc files in the input directory (including sub‑folders) to multi‑page TIFF files.
    // The conversion runs in parallel to improve throughput.
    static void ConvertDocsToTiff(string inputFolder, string outputFolder, int maxDegreeOfParallelism = -1)
    {
        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Gather all .doc files (both .doc and .docx) to process.
        string[] docFiles = Directory.GetFiles(inputFolder, "*.doc", SearchOption.AllDirectories);

        var parallelOptions = new ParallelOptions();
        if (maxDegreeOfParallelism > 0)
            parallelOptions.MaxDegreeOfParallelism = maxDegreeOfParallelism;

        // Process each document concurrently.
        Parallel.ForEach(docFiles, parallelOptions, docPath =>
        {
            try
            {
                // Load the source document using Aspose.Words.
                Document doc = new Document(docPath);

                // Configure image save options for TIFF output.
                ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
                {
                    // Render all pages into a single multi‑frame TIFF.
                    PageLayout = MultiPageLayout.TiffFrames(),
                    // Use LZW compression to keep file size reasonable.
                    TiffCompression = TiffCompression.Lzw,
                    // Set a suitable resolution (e.g., 300 DPI).
                    Resolution = 300
                };

                // Build the output file name.
                string outputFile = Path.Combine(
                    outputFolder,
                    Path.GetFileNameWithoutExtension(docPath) + ".tiff");

                // Save the document as a TIFF using the configured options.
                doc.Save(outputFile, tiffOptions);
            }
            catch (Exception ex)
            {
                // Log conversion errors without stopping the whole batch.
                Console.Error.WriteLine($"Error converting '{docPath}': {ex.Message}");
            }
        });
    }

    // Example entry point.
    static void Main(string[] args)
    {
        // Example usage:
        // args[0] = input folder containing DOC files
        // args[1] = output folder for TIFF files
        // args[2] (optional) = max degree of parallelism
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: BatchDocToTiffConverter <inputFolder> <outputFolder> [maxDegree]");
            return;
        }

        string inputFolder = args[0];
        string outputFolder = args[1];
        int maxDegree = -1;
        if (args.Length >= 3 && int.TryParse(args[2], out int parsed))
            maxDegree = parsed;

        ConvertDocsToTiff(inputFolder, outputFolder, maxDegree);
    }
}

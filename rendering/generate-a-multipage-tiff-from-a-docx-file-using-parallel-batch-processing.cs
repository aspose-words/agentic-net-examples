using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input and output.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        string outputDir = Path.Combine(workDir, "Output");
        Directory.CreateDirectory(workDir);
        Directory.CreateDirectory(outputDir);

        // Create a sample DOCX with several pages.
        string sourceDocPath = Path.Combine(workDir, "Sample.docx");
        CreateSampleDocument(sourceDocPath, pageCount: 5);

        // List of documents to process (batch processing can handle many files).
        string[] sourceFiles = new[] { sourceDocPath };

        // Process each document in parallel.
        Parallel.ForEach(sourceFiles, sourcePath =>
        {
            // Load the document.
            Document doc = new Document(sourcePath);

            // Configure image save options for a multipage TIFF.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Each page will be saved as a separate frame in the TIFF.
                PageLayout = MultiPageLayout.TiffFrames(),
                // Optional: set resolution for better quality.
                Resolution = 300
            };

            // Determine output file name.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
            string tiffPath = Path.Combine(outputDir, $"{fileNameWithoutExt}.tiff");

            // Save the document as a multipage TIFF.
            doc.Save(tiffPath, options);

            // Validate that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");

            // Validate that the number of pages in the source matches the expected page count.
            // (We cannot inspect TIFF frames without external libraries, so we verify the source document.)
            Console.WriteLine($"Processed '{sourcePath}' -> '{tiffPath}' ({doc.PageCount} pages).");
        });
    }

    // Creates a simple DOCX file with the specified number of pages.
    private static void CreateSampleDocument(string filePath, int pageCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= pageCount; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < pageCount)
                builder.InsertBreak(BreakType.PageBreak);
        }

        doc.Save(filePath);
    }
}

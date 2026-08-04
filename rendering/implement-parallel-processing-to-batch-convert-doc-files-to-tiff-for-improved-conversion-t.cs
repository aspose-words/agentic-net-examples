using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare folders for source DOCX files and resulting TIFF files.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(dataDir);
        Directory.CreateDirectory(outputDir);

        // Create a few sample DOCX documents locally.
        const int sampleCount = 5;
        for (int i = 1; i <= sampleCount; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Sample document {i} - Page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Sample document {i} - Page 2.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Sample document {i} - Page 3.");

            string docPath = Path.Combine(dataDir, $"Sample{i}.docx");
            doc.Save(docPath);
        }

        // Gather all DOCX files that need to be converted.
        string[] sourceFiles = Directory.GetFiles(dataDir, "*.docx");

        // Convert each document to a multipage TIFF in parallel.
        Parallel.ForEach(sourceFiles, sourceFile =>
        {
            // Load the source document.
            Document doc = new Document(sourceFile);

            // Determine the output TIFF file path.
            string tiffPath = Path.Combine(outputDir,
                Path.GetFileNameWithoutExtension(sourceFile) + ".tiff");

            // Save the document as TIFF. Each page becomes a frame in the TIFF.
            doc.Save(tiffPath, SaveFormat.Tiff);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF: {tiffPath}");
        });

        // Optional: report the number of files processed.
        Console.WriteLine($"Converted {sourceFiles.Length} documents to TIFF format.");
    }
}

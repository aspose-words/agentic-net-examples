using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Set up folders for input DOCX files and output TIFF files.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample DOCX files (the task must not rely on external files).
        for (int i = 1; i <= 2; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);

            builder.Writeln($"Sample document {i} - Page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Sample document {i} - Page 2.");

            string samplePath = Path.Combine(inputDir, $"Sample{i}.docx");
            sampleDoc.Save(samplePath);
        }

        // Batch convert each DOCX file to a multipage TIFF with custom DPI.
        string[] docxFiles = Directory.GetFiles(inputDir, "*.docx");
        foreach (string docxPath in docxFiles)
        {
            // Load the source document.
            Document doc = new Document(docxPath);

            // Configure image save options for TIFF.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Custom DPI (e.g., 300).
                Resolution = 300,
                // Render all pages into a single multi‑frame TIFF.
                PageLayout = MultiPageLayout.TiffFrames()
            };

            // Determine output file name.
            string tiffPath = Path.Combine(outputDir,
                Path.GetFileNameWithoutExtension(docxPath) + ".tiff");

            // Save the document as a TIFF image.
            doc.Save(tiffPath, options);

            // Validate that the output file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
        }

        // Indicate successful completion (no interactive prompts).
        Console.WriteLine("Batch conversion completed successfully.");
    }
}

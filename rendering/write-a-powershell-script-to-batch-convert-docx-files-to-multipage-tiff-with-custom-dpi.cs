using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample DOCX files.
        for (int i = 1; i <= 3; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Document {i} - Page 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Document {i} - Page 2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Document {i} - Page 3");

            string docPath = Path.Combine(inputDir, $"Sample{i}.docx");
            doc.Save(docPath, SaveFormat.Docx);
        }

        // Batch convert each DOCX to a multipage TIFF with custom DPI.
        foreach (string docxPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docxPath);

            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Set desired DPI (e.g., 300).
                Resolution = 300,
                // Render all pages into a single multi‑frame TIFF.
                PageLayout = MultiPageLayout.TiffFrames()
            };

            string tiffFileName = Path.GetFileNameWithoutExtension(docxPath) + ".tiff";
            string tiffPath = Path.Combine(outputDir, tiffFileName);
            doc.Save(tiffPath, options);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
        }

        // Indicate successful completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }
}

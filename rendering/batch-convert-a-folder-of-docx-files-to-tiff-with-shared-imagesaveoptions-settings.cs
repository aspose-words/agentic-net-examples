using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a folder for sample DOCX files.
        string sourceFolder = Path.Combine(Directory.GetCurrentDirectory(), "SourceDocs");
        Directory.CreateDirectory(sourceFolder);

        // Generate a few sample DOCX documents.
        for (int i = 1; i <= 3; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Document {i} - Page 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Document {i} - Page 2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Document {i} - Page 3");

            string docPath = Path.Combine(sourceFolder, $"Sample{i}.docx");
            doc.Save(docPath);
        }

        // Create a folder for the TIFF output files.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "TiffOutput");
        Directory.CreateDirectory(outputFolder);

        // Shared ImageSaveOptions for all conversions.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render all pages of each document into a single multi‑page TIFF.
            PageLayout = MultiPageLayout.TiffFrames(),
            // Example settings – 300 DPI and LZW compression.
            Resolution = 300,
            TiffCompression = TiffCompression.Lzw
        };

        // Batch convert each DOCX in the source folder to a TIFF file.
        foreach (string docxPath in Directory.GetFiles(sourceFolder, "*.docx"))
        {
            Document doc = new Document(docxPath);

            string tiffPath = Path.Combine(
                outputFolder,
                Path.GetFileNameWithoutExtension(docxPath) + ".tiff");

            doc.Save(tiffPath, tiffOptions);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF: {tiffPath}");
        }

        // Optional: indicate successful completion (no interactive output required).
        Console.WriteLine("Batch conversion completed successfully.");
    }
}

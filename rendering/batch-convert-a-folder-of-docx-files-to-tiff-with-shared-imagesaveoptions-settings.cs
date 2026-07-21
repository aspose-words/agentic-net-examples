using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");

        // Ensure folders exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a few sample DOCX files.
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(inputDir, $"Sample{i}.docx");
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Sample document {i} - Page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Sample document {i} - Page 2.");
            doc.Save(docPath);
        }

        // Shared image save options for TIFF conversion.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = 300,                 // 300 DPI.
            TiffCompression = TiffCompression.Lzw
        };

        // Convert each DOCX in the input folder to a multipage TIFF.
        foreach (string docxFile in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docxFile);
            string tiffFile = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(docxFile) + ".tiff");
            doc.Save(tiffFile, tiffOptions);
        }

        // Verify that all TIFF files were created.
        foreach (string tiffFile in Directory.GetFiles(outputDir, "*.tiff"))
        {
            if (!File.Exists(tiffFile))
                throw new FileNotFoundException("Failed to create TIFF file.", tiffFile);
        }

        Console.WriteLine("Batch conversion completed successfully.");
    }
}

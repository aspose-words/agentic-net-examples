using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define source and output folders inside the system temporary directory.
        string baseDir = Path.Combine(Path.GetTempPath(), "AsposeDemo", "Data");
        string outputDir = Path.Combine(Path.GetTempPath(), "AsposeDemo", "Output");

        // Ensure the folders exist.
        Directory.CreateDirectory(baseDir);
        Directory.CreateDirectory(outputDir);

        // Create a few sample DOCX documents.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample document {i}");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Second page of document {i}");

            string docxPath = Path.Combine(baseDir, $"Sample{i}.docx");
            sampleDoc.Save(docxPath, SaveFormat.Docx);
        }

        // Prepare shared ImageSaveOptions for TIFF conversion.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = 300,                 // 300 DPI
            TiffCompression = TiffCompression.Lzw // LZW compression
        };

        // Convert each DOCX file in the source folder to a multi‑page TIFF.
        string[] docxFiles = Directory.GetFiles(baseDir, "*.docx");
        foreach (string docxFile in docxFiles)
        {
            // Load the document. No password is required for these samples.
            Document doc = new Document(docxFile);

            string tiffFileName = Path.GetFileNameWithoutExtension(docxFile) + ".tiff";
            string tiffPath = Path.Combine(outputDir, tiffFileName);

            // Save the document as a multi‑page TIFF using the shared options.
            doc.Save(tiffPath, tiffOptions);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
        }

        // All conversions completed successfully.
        Console.WriteLine("Batch conversion completed. TIFF files are located in:");
        Console.WriteLine(outputDir);
    }
}

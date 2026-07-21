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

        // Ensure input and output folders exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Create a few sample DOCX files if the input folder is empty.
        // -----------------------------------------------------------------
        if (Directory.GetFiles(inputDir, "*.docx").Length == 0)
        {
            for (int i = 1; i <= 3; i++)
            {
                Document sampleDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(sampleDoc);
                builder.Writeln($"Sample document {i}");
                builder.Writeln("This document will be rendered to a TIFF image.");
                string samplePath = Path.Combine(inputDir, $"Sample{i}.docx");
                sampleDoc.Save(samplePath);
            }
        }

        // -----------------------------------------------------------------
        // Process each DOCX file: render to TIFF with 200 DPI and CCITT3 compression.
        // -----------------------------------------------------------------
        string[] docxFiles = Directory.GetFiles(inputDir, "*.docx");
        foreach (string docxPath in docxFiles)
        {
            // Load the source document.
            Document doc = new Document(docxPath);

            // Configure image save options for TIFF output.
            ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                Resolution = 200,                     // 200 DPI.
                TiffCompression = TiffCompression.Ccitt3 // CCITT3 compression.
            };

            // Determine output file name (same base name, .tiff extension).
            string outputFileName = Path.GetFileNameWithoutExtension(docxPath) + ".tiff";
            string outputPath = Path.Combine(outputDir, outputFileName);

            // Save the document as a TIFF image.
            doc.Save(outputPath, tiffOptions);

            // Verify that the TIFF file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {outputPath}");
        }

        // Optional: indicate completion (no interactive input required).
        Console.WriteLine("Batch processing completed successfully.");
    }
}

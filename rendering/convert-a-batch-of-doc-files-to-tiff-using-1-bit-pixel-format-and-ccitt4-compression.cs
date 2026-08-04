using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Directories for input DOCX files and output TIFF files.
        string inputDir = "InputDocs";
        string outputDir = "OutputTiffs";

        // Ensure the directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a few sample DOCX documents.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample document {i}");
            builder.Writeln("This document will be rendered to a 1‑bit TIFF image using CCITT4 compression.");
            string samplePath = Path.Combine(inputDir, $"Sample{i}.docx");
            sampleDoc.Save(samplePath);
        }

        // Process each DOCX file in the input directory.
        foreach (string docPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the source document.
            Document doc = new Document(docPath);

            // Configure image save options for TIFF output.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Use CCITT4 compression (suitable for 1‑bit images).
                TiffCompression = TiffCompression.Ccitt4,
                // Render the image as 1‑bit indexed (black and white).
                PixelFormat = ImagePixelFormat.Format1bppIndexed,
                // Render all pages into a single multi‑frame TIFF.
                PageLayout = MultiPageLayout.TiffFrames()
            };

            // Determine the output file name.
            string outputFileName = Path.GetFileNameWithoutExtension(docPath) + ".tiff";
            string outputPath = Path.Combine(outputDir, outputFileName);

            // Save the document as a TIFF image.
            doc.Save(outputPath, options);

            // Verify that the TIFF file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {outputPath}");
        }

        // Indicate successful completion.
        Console.WriteLine("Batch conversion to TIFF completed successfully.");
    }
}

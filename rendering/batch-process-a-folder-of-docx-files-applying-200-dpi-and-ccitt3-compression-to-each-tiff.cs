using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing; // Needed for ShapeType enum

public class Program
{
    public static void Main()
    {
        // Define input and output directories relative to the current working directory.
        string inputDir = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "OutputTiffs");

        // Ensure the directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // If the input folder is empty, create a few sample DOCX files.
        if (Directory.GetFiles(inputDir, "*.docx").Length == 0)
        {
            for (int i = 1; i <= 3; i++)
            {
                Document sampleDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(sampleDoc);
                builder.Writeln($"Sample document {i}");
                builder.Writeln("This document will be rendered to a TIFF image with 200 DPI and CCITT3 compression.");
                // Add a simple shape to have some content.
                builder.InsertShape(ShapeType.Rectangle, 100, 50);
                string samplePath = Path.Combine(inputDir, $"Sample{i}.docx");
                sampleDoc.Save(samplePath);
            }
        }

        // Process each DOCX file in the input folder.
        string[] docxFiles = Directory.GetFiles(inputDir, "*.docx");
        foreach (string docxPath in docxFiles)
        {
            // Load the Word document.
            Document doc = new Document(docxPath);

            // Configure image save options for TIFF output.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                Resolution = 200,                     // 200 DPI for both dimensions.
                TiffCompression = TiffCompression.Ccitt3
            };

            // Determine the output TIFF file path.
            string tiffFileName = Path.GetFileNameWithoutExtension(docxPath) + ".tiff";
            string tiffPath = Path.Combine(outputDir, tiffFileName);

            // Save the document as a TIFF image using the specified options.
            doc.Save(tiffPath, options);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
        }

        // Indicate completion.
        Console.WriteLine("Batch processing completed successfully.");
    }
}

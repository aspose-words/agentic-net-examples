using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define base, input and output directories.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");

        // Ensure directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample DOCX files in the input folder.
        for (int i = 1; i <= 2; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"This is sample document {i}.");
            builder.Writeln("It will be rendered to a TIFF image with 200 DPI and CCITT3 compression.");
            string samplePath = Path.Combine(inputDir, $"Sample{i}.docx");
            sampleDoc.Save(samplePath);
        }

        // Process each DOCX file: render to TIFF with required settings.
        foreach (string docxPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the source document.
            Document doc = new Document(docxPath);

            // Configure image save options for TIFF.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                Resolution = 200,                     // Set DPI to 200.
                TiffCompression = TiffCompression.Ccitt3 // Apply CCITT3 compression.
            };

            // Determine output TIFF path.
            string tiffFileName = Path.GetFileNameWithoutExtension(docxPath) + ".tiff";
            string tiffPath = Path.Combine(outputDir, tiffFileName);

            // Save the document as a TIFF image.
            doc.Save(tiffPath, options);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
        }

        // Optional: report the number of processed files.
        int count = Directory.GetFiles(outputDir, "*.tiff").Length;
        Console.WriteLine($"Processed {count} document(s). TIFF files are located in: {outputDir}");
    }
}

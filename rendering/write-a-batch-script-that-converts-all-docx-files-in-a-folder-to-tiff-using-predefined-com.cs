using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

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

        // Create a few sample DOCX files for demonstration purposes.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"This is sample document #{i}.");
            builder.Writeln("The quick brown fox jumps over the lazy dog.");
            string samplePath = Path.Combine(inputDir, $"Sample{i}.docx");
            sampleDoc.Save(samplePath, SaveFormat.Docx);
        }

        // Process each DOCX file in the input directory.
        foreach (string docxPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the DOCX document.
            Document doc = new Document(docxPath);

            // Configure image save options for TIFF with predefined compression.
            ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Use LZW compression (default is LZW, but set explicitly for clarity).
                TiffCompression = TiffCompression.Lzw,
                // Optional: set resolution to 300 DPI.
                Resolution = 300
            };

            // Determine the output TIFF file path.
            string tiffFileName = Path.GetFileNameWithoutExtension(docxPath) + ".tiff";
            string tiffPath = Path.Combine(outputDir, tiffFileName);

            // Save the document as a multipage TIFF.
            doc.Save(tiffPath, tiffOptions);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
            {
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
            }
        }

        // Indicate completion (no interactive prompts).
        Console.WriteLine("Conversion of all DOCX files to TIFF completed successfully.");
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define input and output directories.
        string inputDir = "InputDocs";
        string outputDir = "OutputTiffs";

        // Ensure the directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // If there are no DOCX files, create a few sample documents.
        if (Directory.GetFiles(inputDir, "*.docx").Length == 0)
        {
            for (int i = 1; i <= 3; i++)
            {
                Document sampleDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(sampleDoc);
                builder.Writeln($"Sample document {i} - Page 1.");
                builder.InsertBreak(BreakType.PageBreak);
                builder.Writeln($"Sample document {i} - Page 2.");
                string samplePath = Path.Combine(inputDir, $"Sample{i}.docx");
                sampleDoc.Save(samplePath);
            }
        }

        // Process each DOCX file in the input folder.
        foreach (string docxPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the DOCX document.
            Document doc = new Document(docxPath);

            // Configure TIFF save options with predefined compression.
            ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                TiffCompression = TiffCompression.Lzw // Use LZW compression.
            };

            // Determine the output TIFF file path.
            string tiffPath = Path.Combine(outputDir,
                Path.GetFileNameWithoutExtension(docxPath) + ".tiff");

            // Save the document as a multipage TIFF.
            doc.Save(tiffPath, tiffOptions);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF for '{docxPath}'.");
        }

        Console.WriteLine("All DOCX files have been converted to TIFF.");
    }
}

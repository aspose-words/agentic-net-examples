using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a temporary folder for input and output files.
        string baseDir = Path.Combine(Path.GetTempPath(), "AsposeExample");
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Generate sample DOCX files if they do not already exist.
        var docxFiles = new List<string>();
        for (int i = 1; i <= 2; i++)
        {
            string filePath = Path.Combine(inputDir, $"Document{i}.docx");
            if (!File.Exists(filePath))
            {
                var doc = new Document();
                var builder = new DocumentBuilder(doc);
                builder.Writeln($"This is sample content for Document {i}.");
                doc.Save(filePath);
            }
            docxFiles.Add(filePath);
        }

        // Process each DOCX file in parallel and save as a multi‑page TIFF.
        Parallel.ForEach(docxFiles, docxPath =>
        {
            // Load the DOCX document.
            var doc = new Document(docxPath);

            // Configure image save options for TIFF format.
            var options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Each page becomes a separate frame in the resulting multi‑frame TIFF.
                PageLayout = MultiPageLayout.TiffFrames()
            };

            // Build the output file name.
            string fileName = Path.GetFileNameWithoutExtension(docxPath) + ".tiff";
            string outputPath = Path.Combine(outputDir, fileName);

            // Save the document as a multi‑page TIFF.
            doc.Save(outputPath, options);
        });

        Console.WriteLine($"Conversion completed. TIFF files are located in: {outputDir}");
    }
}

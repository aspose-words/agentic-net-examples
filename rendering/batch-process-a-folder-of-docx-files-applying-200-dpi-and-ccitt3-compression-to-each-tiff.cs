using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for source DOCX files and generated TIFF files.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "DemoData");
        string sourceDir = Path.Combine(baseDir, "SourceDocs");
        string outputDir = Path.Combine(baseDir, "TiffOutputs");

        // Ensure the folders exist.
        Directory.CreateDirectory(sourceDir);
        Directory.CreateDirectory(outputDir);

        // Create a few sample DOCX files if the source folder is empty.
        if (Directory.GetFiles(sourceDir, "*.docx").Length == 0)
        {
            for (int i = 1; i <= 3; i++)
            {
                string docPath = Path.Combine(sourceDir, $"Sample{i}.docx");
                Document sampleDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(sampleDoc);
                builder.Writeln($"This is sample document #{i}.");
                builder.InsertBreak(BreakType.PageBreak);
                builder.Writeln($"Second page of sample document #{i}.");
                sampleDoc.Save(docPath);
            }
        }

        // Process each DOCX file: render to a single multi‑page TIFF with 200 DPI and CCITT3 compression.
        foreach (string docxPath in Directory.GetFiles(sourceDir, "*.docx"))
        {
            // Load the Word document.
            Document doc = new Document(docxPath);

            // Configure image save options for TIFF output.
            ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                Resolution = 200,                     // 200 DPI.
                TiffCompression = TiffCompression.Ccitt3
            };

            // Determine the output TIFF file name.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(docxPath);
            string tiffPath = Path.Combine(outputDir, $"{fileNameWithoutExt}.tiff");

            // Save the document as a TIFF image.
            doc.Save(tiffPath, tiffOptions);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
        }

        // All processing completed successfully.
        Console.WriteLine("Batch conversion completed. TIFF files are located at:");
        Console.WriteLine(outputDir);
    }
}

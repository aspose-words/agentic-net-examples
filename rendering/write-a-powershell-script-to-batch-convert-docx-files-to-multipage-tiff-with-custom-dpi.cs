using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define temporary folders for input DOCX files and output TIFF files.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "BatchConversionDemo");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputTiffs");

        // Ensure clean environment.
        if (Directory.Exists(baseDir))
            Directory.Delete(baseDir, true);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a few sample DOCX documents with multiple pages.
        for (int i = 1; i <= 3; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Document {i} - Page 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Document {i} - Page 2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Document {i} - Page 3");

            string docPath = Path.Combine(inputDir, $"Sample{i}.docx");
            doc.Save(docPath);
        }

        // Custom DPI for the TIFF images.
        const float customDpi = 300f;

        // Process each DOCX file in the input folder.
        foreach (string docPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the source document.
            Document sourceDoc = new Document(docPath);

            // Configure image save options for multipage TIFF.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Set the desired resolution (DPI).
                Resolution = customDpi,
                // Render each page as a separate frame in the TIFF.
                PageLayout = MultiPageLayout.TiffFrames()
            };

            // Determine output file name.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(docPath);
            string tiffPath = Path.Combine(outputDir, $"{fileNameWithoutExt}.tiff");

            // Save the document as a multipage TIFF.
            sourceDoc.Save(tiffPath, saveOptions);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
        }

        // Optional: indicate successful completion (no interactive I/O required).
        Console.WriteLine("Batch conversion completed successfully.");
    }
}

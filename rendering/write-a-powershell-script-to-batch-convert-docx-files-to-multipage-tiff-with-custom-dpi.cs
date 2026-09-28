using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class BatchDocxToTiff
{
    public static void Main()
    {
        // Define temporary base folder and subfolders for source DOCX files and output TIFF files.
        string baseFolder = Path.Combine(Path.GetTempPath(), "AsposeBatchConvert");
        string sourceFolder = Path.Combine(baseFolder, "Source");
        string outputFolder = Path.Combine(baseFolder, "Output");

        // Ensure a clean environment.
        if (Directory.Exists(baseFolder))
            Directory.Delete(baseFolder, true);
        Directory.CreateDirectory(sourceFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample DOCX files with multiple pages.
        CreateSampleDocx(Path.Combine(sourceFolder, "Sample1.docx"), "First document", 3);
        CreateSampleDocx(Path.Combine(sourceFolder, "Sample2.docx"), "Second document", 5);

        // Desired DPI for the TIFF images.
        const int dpi = 300;

        // Process each DOCX file in the source folder.
        foreach (string docxPath in Directory.GetFiles(sourceFolder, "*.docx"))
        {
            // Load the DOCX document.
            Document doc = new Document(docxPath);

            // Configure TIFF save options.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Set both horizontal and vertical resolution.
                Resolution = dpi
                // By default all pages are rendered; no need to set PageSet.
            };

            // Build the output TIFF file path.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(docxPath);
            string tiffPath = Path.Combine(outputFolder, fileNameWithoutExt + ".tiff");

            // Save the document as a multipage TIFF.
            doc.Save(tiffPath, saveOptions);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
        }

        // Indicate successful batch conversion.
        Console.WriteLine("Batch conversion completed successfully.");
    }

    // Helper method that creates a DOCX file containing the specified number of pages.
    private static void CreateSampleDocx(string filePath, string title, int pageCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln(title);
        builder.Writeln($"This document contains {pageCount} pages.");

        // Insert page breaks to generate additional pages.
        for (int i = 1; i < pageCount; i++)
        {
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Page {i + 1}");
        }

        doc.Save(filePath);
    }
}

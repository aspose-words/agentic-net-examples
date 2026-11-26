using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a temporary working folder.
        string workFolder = Path.Combine(Path.GetTempPath(), "DocxToTiffDemo");
        Directory.CreateDirectory(workFolder);

        // Create sample DOCX files.
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(workFolder, $"Sample{i}.docx");
            CreateSampleDocument(docPath, i * 2); // Vary page count.
        }

        // Define custom DPI.
        const int customDpi = 300;

        // Process each DOCX file in the folder.
        foreach (string docxFile in Directory.GetFiles(workFolder, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(docxFile);

            // Configure image save options for multipage TIFF.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Set the desired DPI.
                Resolution = customDpi,
                // Ensure all pages are saved to the same TIFF file.
                PageSet = PageSet.All,
                // Use CCITT Group 4 compression (available in the current API).
                TiffCompression = TiffCompression.Ccitt4
            };

            // Determine output TIFF path.
            string tiffPath = Path.ChangeExtension(docxFile, ".tiff");

            // Save the document as a multipage TIFF.
            doc.Save(tiffPath, options);

            // Validate that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
        }

        // Optional cleanup: delete the temporary folder.
        // Directory.Delete(workFolder, true);
    }

    // Helper method to create a sample DOCX with a given number of pages.
    private static void CreateSampleDocument(string filePath, int pageCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int p = 1; p <= pageCount; p++)
        {
            builder.Writeln($"Page {p}");
            // Add enough text to force a new page.
            for (int i = 0; i < 30; i++)
                builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit.");

            if (p < pageCount)
                builder.InsertBreak(BreakType.PageBreak);
        }

        doc.Save(filePath);
    }
}

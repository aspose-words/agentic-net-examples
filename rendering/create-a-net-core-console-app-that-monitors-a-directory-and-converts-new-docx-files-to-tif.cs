using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder that will be watched.
        string watchFolder = Path.Combine(Path.GetTempPath(), "DocxWatchDemo");
        Directory.CreateDirectory(watchFolder);

        // Create a sample DOCX file in the folder (simulating a new file arriving).
        string sampleDocPath = Path.Combine(watchFolder, "SampleDocument.docx");
        CreateSampleDocument(sampleDocPath);

        // Process each DOCX file found in the folder.
        foreach (string docxFile in Directory.GetFiles(watchFolder, "*.docx"))
        {
            // Load the DOCX document.
            Document doc = new Document(docxFile);

            // Configure TIFF save options – render all pages into a single multipage TIFF.
            ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Render every page of the source document.
                PageSet = PageSet.All
            };

            // Determine the output TIFF path.
            string tiffPath = Path.ChangeExtension(docxFile, ".tiff");

            // Save the document as a multipage TIFF.
            doc.Save(tiffPath, tiffOptions);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");

            // Move the processed DOCX to a subfolder to avoid re‑processing.
            string processedFolder = Path.Combine(watchFolder, "Processed");
            Directory.CreateDirectory(processedFolder);
            string destDocxPath = Path.Combine(processedFolder, Path.GetFileName(docxFile));
            File.Move(docxFile, destDocxPath);
        }

        // Cleanup: delete the temporary folder and its contents.
        // Comment out the following line if you wish to inspect the files after execution.
        Directory.Delete(watchFolder, true);
    }

    // Helper method that creates a simple three‑page DOCX document.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"This is page {i} of the sample document.");
            if (i < 3)
                builder.InsertBreak(BreakType.PageBreak);
        }

        doc.Save(filePath);
    }
}

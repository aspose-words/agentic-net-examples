using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for sample DOCX files.
        string inputFolder = Path.Combine(Path.GetTempPath(), "AsposeWordsSampleInput");
        Directory.CreateDirectory(inputFolder);

        // Create a few sample DOCX documents.
        for (int i = 1; i <= 2; i++)
        {
            string docPath = Path.Combine(inputFolder, $"Sample{i}.docx");
            Document doc = new Document();
            // Add a simple paragraph with text.
            doc.FirstSection.Body.FirstParagraph.AppendChild(new Run(doc, $"This is sample document {i}."));
            doc.Save(docPath);
        }

        // Process each DOCX file in the folder, rendering to TIFF with 200 DPI and CCITT3 compression.
        foreach (string docxFile in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document document = new Document(docxFile);

            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                Resolution = 200,                     // Set DPI to 200.
                TiffCompression = TiffCompression.Ccitt3 // Apply CCITT3 compression.
            };

            string tiffFile = Path.ChangeExtension(docxFile, ".tiff");
            document.Save(tiffFile, options);

            // Validate that the TIFF file was created.
            if (!File.Exists(tiffFile))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffFile}");
        }

        // Optional: clean up temporary files (comment out if you want to inspect the results).
        // Directory.Delete(inputFolder, true);
    }
}

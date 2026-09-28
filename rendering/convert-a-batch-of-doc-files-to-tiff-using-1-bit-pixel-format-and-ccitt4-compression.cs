using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for the sample documents and output TIFF files.
        string workFolder = Path.Combine(Path.GetTempPath(), "AsposeWordsTiffBatch");
        Directory.CreateDirectory(workFolder);

        // Create a few sample DOCX files.
        CreateSampleDocument(Path.Combine(workFolder, "Sample1.docx"), "This is the first sample document.");
        CreateSampleDocument(Path.Combine(workFolder, "Sample2.docx"), "This is the second sample document.");

        // Find all DOC and DOCX files in the folder.
        string[] docFiles = Directory.GetFiles(workFolder, "*.doc*");

        foreach (string docPath in docFiles)
        {
            // Load the source document.
            Document doc = new Document(docPath);

            // Configure TIFF save options: 1‑bit (black‑and‑white) and CCITT4 compression.
            ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                ImageColorMode = ImageColorMode.BlackAndWhite,
                TiffCompression = TiffCompression.Ccitt4
            };

            // Determine the output TIFF file path.
            string tiffPath = Path.ChangeExtension(docPath, ".tiff");

            // Save the document as a multipage TIFF.
            doc.Save(tiffPath, tiffOptions);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
        }

        // Optional: indicate successful completion (no interactive prompts).
        Console.WriteLine("Batch conversion completed successfully.");
    }

    // Helper method to create a simple DOCX document with supplied text.
    private static void CreateSampleDocument(string filePath, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(text);
        doc.Save(filePath);
    }
}

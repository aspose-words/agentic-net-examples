using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for sample DOCX files.
        string sourceFolder = Path.Combine(Directory.GetCurrentDirectory(), "BatchDocs");
        Directory.CreateDirectory(sourceFolder);

        // Generate a few sample DOCX documents.
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(sourceFolder, $"Sample{i}.docx");
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document {i}.");
            doc.Save(docPath);
        }

        // Create an output folder for the generated TIFF files.
        string outputFolder = Path.Combine(sourceFolder, "Output");
        Directory.CreateDirectory(outputFolder);

        // Process each DOCX file in the source folder.
        string[] docFiles = Directory.GetFiles(sourceFolder, "*.docx");
        foreach (string docFile in docFiles)
        {
            // Load the source document.
            Document document = new Document(docFile);

            // Configure TIFF rendering options: 1‑bit (black‑and‑white) and CCITT4 compression.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                ImageColorMode = ImageColorMode.BlackAndWhite,
                TiffCompression = TiffCompression.Ccitt4,
                PageSet = PageSet.All // Render all pages.
            };

            // Determine the output TIFF file path.
            string tiffPath = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(docFile) + ".tiff");

            // Save the document as a multipage TIFF.
            document.Save(tiffPath, saveOptions);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new Exception($"Failed to create TIFF file: {tiffPath}");
        }

        // Optional: confirm that the number of generated TIFF files matches the source count.
        int expectedCount = docFiles.Length;
        int actualCount = Directory.GetFiles(outputFolder, "*.tiff").Length;
        if (actualCount != expectedCount)
            throw new Exception($"Expected {expectedCount} TIFF files, but found {actualCount}.");

        // Program completed successfully.
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder to hold sample DOCX files and the resulting TIFFs.
        string docsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        Directory.CreateDirectory(docsFolder);

        // Create a few sample DOCX files locally.
        CreateSampleDocument(Path.Combine(docsFolder, "Sample1.docx"), "First sample document.");
        CreateSampleDocument(Path.Combine(docsFolder, "Sample2.docx"), "Second sample document with multiple pages.\nPage break below.\n\n\nPage 2 content.");
        CreateSampleDocument(Path.Combine(docsFolder, "Sample3.docx"), "Third sample document.");

        // Process each DOCX file in the folder.
        foreach (string docxPath in Directory.GetFiles(docsFolder, "*.docx"))
        {
            // Load the Word document.
            Document doc = new Document(docxPath);

            // Configure image save options for TIFF output.
            ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Set resolution to 200 DPI.
                Resolution = 200f,
                // Apply CCITT3 compression.
                TiffCompression = TiffCompression.Ccitt3
            };

            // Determine the output TIFF file path.
            string tiffPath = Path.Combine(
                docsFolder,
                Path.GetFileNameWithoutExtension(docxPath) + ".tiff");

            // Save the document as a multipage TIFF using the specified options.
            doc.Save(tiffPath, tiffOptions);

            // Validate that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
        }

        // Optional: indicate processing is complete.
        Console.WriteLine("Batch conversion completed successfully.");
    }

    // Helper method to create a simple DOCX file with given text.
    private static void CreateSampleDocument(string filePath, string content)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(content);
        doc.Save(filePath);
    }
}

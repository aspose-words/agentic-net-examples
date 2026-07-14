using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define input and output folders relative to the current directory.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Input");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample DOCX files if the input folder is empty.
        if (Directory.GetFiles(inputFolder, "*.docx").Length == 0)
        {
            CreateSampleDocument(Path.Combine(inputFolder, "Sample1.docx"), "Hello Aspose.Words!");
            CreateSampleDocument(Path.Combine(inputFolder, "Sample2.docx"), "Batch conversion to TIFF.");
        }

        // Get all DOCX files in the input folder.
        string[] docxFiles = Directory.GetFiles(inputFolder, "*.docx");

        foreach (string docxPath in docxFiles)
        {
            // Load the DOCX document.
            Document doc = new Document(docxPath);

            // Configure TIFF save options with predefined compression.
            ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                TiffCompression = TiffCompression.Lzw
            };

            // Determine the output TIFF file path.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(docxPath);
            string tiffPath = Path.Combine(outputFolder, fileNameWithoutExt + ".tiff");

            // Save the document as a TIFF image.
            doc.Save(tiffPath, tiffOptions);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
        }
    }

    // Helper method to create a simple DOCX document with specified text.
    private static void CreateSampleDocument(string filePath, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(text);
        doc.Save(filePath);
    }
}

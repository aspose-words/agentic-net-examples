using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define base, input, and output folders.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "BatchConversionDemo");
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string outputFolder = Path.Combine(baseDir, "OutputTiffs");

        // Ensure a clean environment.
        if (Directory.Exists(baseDir))
            Directory.Delete(baseDir, true);
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample DOCX files.
        CreateSampleDocx(Path.Combine(inputFolder, "Sample1.docx"), "First sample document.");
        CreateSampleDocx(
            Path.Combine(inputFolder, "Sample2.docx"),
            "Second sample document with multiple lines.\nLine 2.\nLine 3."
        );

        // Shared ImageSaveOptions for TIFF conversion.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render all pages of each document into a multipage TIFF.
            PageSet = PageSet.All,
            // Set resolution (dpi) and compression.
            Resolution = 300,
            TiffCompression = TiffCompression.Lzw
        };

        // Batch conversion.
        foreach (string docxPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the DOCX document.
            Document doc = new Document(docxPath);

            // Determine output TIFF path.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(docxPath);
            string tiffPath = Path.Combine(outputFolder, fileNameWithoutExt + ".tiff");

            // Save as TIFF using the shared options.
            doc.Save(tiffPath, tiffOptions);

            // Validate that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new Exception($"Failed to create TIFF file: {tiffPath}");
        }

        // Indicate successful completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }

    // Helper method to create a simple DOCX file with given text.
    private static void CreateSampleDocx(string filePath, string content)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(content);
        doc.Save(filePath);
    }
}

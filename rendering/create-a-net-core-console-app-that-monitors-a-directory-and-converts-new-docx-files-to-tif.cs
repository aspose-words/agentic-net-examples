using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define input and output directories relative to the current working directory.
        string inputDir = Path.Combine(Directory.GetCurrentDirectory(), "Input");
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");

        // Ensure the directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a sample DOCX file in the input directory.
        string sampleDocPath = Path.Combine(inputDir, "SampleDocument.docx");
        CreateSampleDocument(sampleDocPath);

        // Process each DOCX file found in the input directory.
        foreach (string docxFile in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the DOCX document.
            Document doc = new Document(docxFile);

            // Prepare the output TIFF file path.
            string tiffFileName = Path.GetFileNameWithoutExtension(docxFile) + ".tiff";
            string tiffFilePath = Path.Combine(outputDir, tiffFileName);

            // Configure image save options for TIFF rendering.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Example: set a higher resolution for better quality.
                Resolution = 300
            };

            // Save the document as a TIFF image (multi‑page if the source has multiple pages).
            doc.Save(tiffFilePath, saveOptions);

            // Validate that the TIFF file was created.
            if (!File.Exists(tiffFilePath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffFilePath}");
        }

        // Indicate successful completion.
        Console.WriteLine("Conversion completed.");
    }

    // Creates a simple DOCX document with some text content.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document generated for conversion to TIFF.");
        builder.Writeln("It contains multiple lines of text to demonstrate rendering.");
        doc.Save(filePath, SaveFormat.Docx);
    }
}

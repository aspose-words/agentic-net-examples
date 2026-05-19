using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define input and output directories relative to the current working directory.
        string inputDir = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "OutputTiffs");

        // Ensure the directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // If there are no DOCX files, create a sample document to demonstrate the conversion.
        if (Directory.GetFiles(inputDir, "*.docx").Length == 0)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("This is a sample DOCX file generated for batch conversion.");
            string samplePath = Path.Combine(inputDir, "Sample.docx");
            sampleDoc.Save(samplePath);
        }

        // Retrieve all DOCX files from the input directory.
        string[] docxFiles = Directory.GetFiles(inputDir, "*.docx");

        // Define the desired TIFF compression (predefined).
        TiffCompression compression = TiffCompression.Lzw;

        foreach (string docxPath in docxFiles)
        {
            // Load the DOCX document.
            Document doc = new Document(docxPath);

            // Configure image save options for TIFF with the specified compression.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                TiffCompression = compression
            };

            // Determine the output TIFF file path.
            string tiffFileName = Path.GetFileNameWithoutExtension(docxPath) + ".tiff";
            string tiffPath = Path.Combine(outputDir, tiffFileName);

            // Save the document as a TIFF image.
            doc.Save(tiffPath, saveOptions);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
        }

        // Optional: indicate completion (no interactive input required).
        Console.WriteLine($"Converted {docxFiles.Length} DOCX file(s) to TIFF in '{outputDir}'.");
    }
}

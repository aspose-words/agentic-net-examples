using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input DOCX files and output TIFF files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputTiffs");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a few sample DOCX files to demonstrate batch conversion.
        CreateSampleDocuments(inputFolder, 3);

        // Configure shared ImageSaveOptions for TIFF conversion.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use LZW compression (default) – can be changed as needed.
            TiffCompression = TiffCompression.Lzw,
            // Render at 300 DPI for decent quality.
            Resolution = 300,
            // Render each page as a separate frame in the multi‑page TIFF.
            PageLayout = MultiPageLayout.TiffFrames()
        };

        // Process each DOCX file in the input folder.
        foreach (string docxPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the source document.
            Document doc = new Document(docxPath);

            // Determine the output TIFF file name.
            string tiffPath = Path.Combine(
                outputFolder,
                Path.GetFileNameWithoutExtension(docxPath) + ".tiff");

            // Save the document as a TIFF using the shared options.
            doc.Save(tiffPath, tiffOptions);
        }

        // Verify that each expected TIFF file was created.
        foreach (string tiffPath in Directory.GetFiles(outputFolder, "*.tiff"))
        {
            if (!File.Exists(tiffPath))
                throw new FileNotFoundException($"Failed to create TIFF file: {tiffPath}");
        }

        // Indicate successful completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }

    // Helper method to create a specified number of simple DOCX files.
    private static void CreateSampleDocuments(string folder, int count)
    {
        for (int i = 1; i <= count; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Sample Document {i}");
            builder.Writeln("This is the first page.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("This is the second page.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("This is the third page.");

            string docxPath = Path.Combine(folder, $"Sample{i}.docx");
            doc.Save(docxPath, SaveFormat.Docx);
        }
    }
}

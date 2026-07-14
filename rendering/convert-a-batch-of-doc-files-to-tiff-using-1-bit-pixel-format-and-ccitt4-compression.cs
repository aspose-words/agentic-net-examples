using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input DOC files and output TIFF files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputTiffs");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a few sample DOC files.
        CreateSampleDocument(Path.Combine(inputFolder, "Sample1.doc"));
        CreateSampleDocument(Path.Combine(inputFolder, "Sample2.doc"));
        CreateSampleDocument(Path.Combine(inputFolder, "Sample3.doc"));

        // Process each DOC file in the input folder.
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.doc"))
        {
            // Load the document.
            Document doc = new Document(docPath);

            // Configure image save options for TIFF output.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Use CCITT4 compression.
                TiffCompression = TiffCompression.Ccitt4,
                // Render the image as 1‑bit (black and white).
                PixelFormat = ImagePixelFormat.Format1bppIndexed,
                // Render each page as a separate frame in a multi‑page TIFF.
                PageLayout = MultiPageLayout.TiffFrames()
            };

            // Determine the output file name.
            string outputFileName = Path.GetFileNameWithoutExtension(docPath) + ".tiff";
            string outputPath = Path.Combine(outputFolder, outputFileName);

            // Save the document as a TIFF file.
            doc.Save(outputPath, options);

            // Verify that the TIFF file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {outputPath}");
        }

        // Optional: indicate successful completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }

    // Helper method to create a simple DOC file with multiple pages.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some content spanning three pages.
        builder.Writeln("This is the first page of the document.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is the second page of the document.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is the third page of the document.");

        // Save the document in DOC format.
        doc.Save(filePath, SaveFormat.Doc);
    }
}

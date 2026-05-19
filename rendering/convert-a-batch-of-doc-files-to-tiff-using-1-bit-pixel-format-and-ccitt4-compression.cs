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

        // Create a few sample DOC files to demonstrate batch processing.
        CreateSampleDocument(Path.Combine(inputFolder, "Sample1.doc"), "First document", 2);
        CreateSampleDocument(Path.Combine(inputFolder, "Sample2.doc"), "Second document", 3);
        CreateSampleDocument(Path.Combine(inputFolder, "Sample3.doc"), "Third document", 1);

        // Process each DOC file in the input folder.
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.doc"))
        {
            // Load the document.
            Document doc = new Document(docPath);

            // Configure image save options for TIFF output.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Render each page as a separate frame in a multi‑page TIFF.
                PageLayout = MultiPageLayout.TiffFrames(),
                // Use 1‑bit per pixel.
                PixelFormat = ImagePixelFormat.Format1bppIndexed,
                // Apply CCITT4 compression (suitable for 1‑bpp images).
                TiffCompression = TiffCompression.Ccitt4,
                // Optional: set a reasonable resolution.
                Resolution = 300
            };

            // Build the output file name.
            string outputFileName = Path.GetFileNameWithoutExtension(docPath) + ".tiff";
            string outputPath = Path.Combine(outputFolder, outputFileName);

            // Save the document as a TIFF image.
            doc.Save(outputPath, options);

            // Verify that the TIFF file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {outputPath}");
        }

        // Indicate successful completion (no interactive output required).
        Console.WriteLine("Batch conversion completed successfully.");
    }

    // Helper method to create a simple DOC file with the specified title and number of pages.
    private static void CreateSampleDocument(string filePath, string title, int pageCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln(title);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

        for (int i = 1; i <= pageCount; i++)
        {
            builder.Writeln($"This is page {i} of the document.");
            if (i < pageCount)
                builder.InsertBreak(BreakType.PageBreak);
        }

        doc.Save(filePath);
    }
}

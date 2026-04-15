using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare input and output directories.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputTiffs");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample DOC files.
        CreateSampleDocument(Path.Combine(inputDir, "Sample1.doc"), "First document", 2);
        CreateSampleDocument(Path.Combine(inputDir, "Sample2.doc"), "Second document", 3);
        CreateSampleDocument(Path.Combine(inputDir, "Sample3.doc"), "Third document", 1);

        // Process each DOC file.
        foreach (string docPath in Directory.GetFiles(inputDir, "*.doc"))
        {
            // Load the source document.
            Document doc = new Document(docPath);

            // Configure TIFF save options: 1‑bpp pixel format and CCITT4 compression.
            ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                TiffCompression = TiffCompression.Ccitt4,
                PixelFormat = ImagePixelFormat.Format1bppIndexed,
                PageLayout = MultiPageLayout.TiffFrames()
            };

            // Determine output file name.
            string outFileName = Path.GetFileNameWithoutExtension(docPath) + ".tiff";
            string outPath = Path.Combine(outputDir, outFileName);

            // Save the document as a multipage TIFF.
            doc.Save(outPath, tiffOptions);

            // Verify that the TIFF file was created.
            if (!File.Exists(outPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {outPath}");
        }

        // Indicate successful completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }

    // Helper method to create a simple DOC file with the specified text and page count.
    private static void CreateSampleDocument(string filePath, string title, int pageCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln(title);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

        for (int i = 1; i <= pageCount; i++)
        {
            builder.Writeln($"Content of page {i}.");
            if (i < pageCount)
                builder.InsertBreak(BreakType.PageBreak);
        }

        doc.Save(filePath);
    }
}

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;

public class BatchDocToTiff
{
    public static void Main()
    {
        // Create a temporary directory for sample DOC files.
        string inputDir = Path.Combine(Path.GetTempPath(), "DocToTiffInput");
        Directory.CreateDirectory(inputDir);

        // Create a temporary directory for output TIFF files.
        string outputDir = Path.Combine(Path.GetTempPath(), "DocToTiffOutput");
        Directory.CreateDirectory(outputDir);

        // Generate sample DOC files.
        List<string> docPaths = new List<string>();
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(inputDir, $"Sample{i}.doc");
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document {i}.");
            // Add a second page to test multi‑page TIFF.
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Second page of document {i}.");
            doc.Save(docPath);
            docPaths.Add(docPath);
        }

        // Convert each DOC file to a multi‑page TIFF with 1‑bit color and CCITT4 compression.
        foreach (string docPath in docPaths)
        {
            Document doc = new Document(docPath);

            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Render all pages into a single multi‑page TIFF.
                PageSet = PageSet.All,
                // Use 1‑bit (black and white) color mode.
                ImageColorMode = ImageColorMode.BlackAndWhite,
                // Set CCITT4 compression.
                TiffCompression = TiffCompression.Ccitt4,
                // Optional: set resolution (dpi) if desired.
                Resolution = 300
            };

            string tiffPath = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(docPath) + ".tiff");
            doc.Save(tiffPath, options);

            // Validate that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
        }

        // Clean up temporary input files (optional).
        // Directory.Delete(inputDir, true);
        // Output directory remains for inspection.
    }
}

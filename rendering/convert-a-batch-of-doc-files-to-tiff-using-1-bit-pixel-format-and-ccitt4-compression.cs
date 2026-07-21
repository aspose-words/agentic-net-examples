using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder for all generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a few sample DOCX files (the task requires DOC files,
        //    but DOCX is also a Word format and can be loaded directly).
        // -----------------------------------------------------------------
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);

            builder.Writeln($"Sample document {i} - Page 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Sample document {i} - Page 2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Sample document {i} - Page 3");

            string docPath = Path.Combine(artifactsDir, $"Sample{i}.docx");
            sampleDoc.Save(docPath);
        }

        // -----------------------------------------------------------------
        // 2. Convert each DOC/DOCX file in the folder to a single TIFF file
        //    using 1‑bit pixel format and CCITT4 compression.
        // -----------------------------------------------------------------
        string[] sourceFiles = Directory.GetFiles(artifactsDir, "*.doc*");

        foreach (string sourcePath in sourceFiles)
        {
            // Load the source document.
            Document doc = new Document(sourcePath);

            // Configure image save options for TIFF output.
            ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Use CCITT4 compression (suitable for 1‑bpp images).
                TiffCompression = TiffCompression.Ccitt4,

                // Force a 1‑bit per pixel format.
                PixelFormat = ImagePixelFormat.Format1bppIndexed,

                // Render all pages into a multi‑frame TIFF.
                PageLayout = MultiPageLayout.TiffFrames()
            };

            // Destination TIFF file path (same name, .tiff extension).
            string tiffPath = Path.ChangeExtension(sourcePath, ".tiff");

            // Save the document as TIFF.
            doc.Save(tiffPath, tiffOptions);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");

            Console.WriteLine($"Converted '{Path.GetFileName(sourcePath)}' to '{Path.GetFileName(tiffPath)}'.");
        }

        Console.WriteLine("Batch conversion completed successfully.");
    }
}

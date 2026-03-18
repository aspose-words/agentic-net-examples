using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class BatchDocToTiff
{
    static void Main()
    {
        // Folder containing source DOC/DOCX files.
        string sourceFolder = @"C:\Docs\Input";
        // Folder where the resulting TIFF files will be saved.
        string outputFolder = @"C:\Docs\Output";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Get all Word documents in the source folder (DOC and DOCX).
        string[] docFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string filePath in docFiles)
        {
            // Process only .doc and .docx files.
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext != ".doc" && ext != ".docx")
                continue;

            // Load the document (lifecycle rule: load).
            Document doc = new Document(filePath);

            // Configure image save options for TIFF:
            // - Use TIFF format.
            // - Apply CCITT4 compression.
            // - Render as 1‑bit per pixel (indexed).
            // - Use a binarization method suitable for CCITT compression.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                TiffCompression = TiffCompression.Ccitt4,
                PixelFormat = ImagePixelFormat.Format1bppIndexed,
                TiffBinarizationMethod = ImageBinarizationMethod.Threshold,
                // Render each page as a separate frame in a multi‑page TIFF.
                PageLayout = MultiPageLayout.TiffFrames()
            };

            // Build the output file name (same base name, .tiff extension).
            string outputFile = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(filePath) + ".tiff");

            // Save the document using the configured options (lifecycle rule: save).
            doc.Save(outputFile, options);
        }
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class BatchDocxToTiff
{
    static void Main()
    {
        // Folder containing the source DOCX files.
        string sourceFolder = @"C:\Docs\Input";
        // Folder where the resulting TIFF files will be saved.
        string outputFolder = @"C:\Docs\Output";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Shared ImageSaveOptions for all conversions.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use LZW compression (default) – change as needed.
            TiffCompression = TiffCompression.Lzw,
            // Set desired resolution (dots per inch).
            Resolution = 300,
            // Render all pages into a single multi‑frame TIFF.
            PageLayout = MultiPageLayout.TiffFrames()
        };

        // Process each DOCX file in the source folder.
        foreach (string docxPath in Directory.GetFiles(sourceFolder, "*.docx"))
        {
            // Load the DOCX document.
            Document doc = new Document(docxPath);

            // Build the output TIFF file name (same base name, .tiff extension).
            string tiffPath = Path.Combine(outputFolder,
                Path.GetFileNameWithoutExtension(docxPath) + ".tiff");

            // Save the document as a TIFF using the shared options.
            doc.Save(tiffPath, tiffOptions);
        }
    }
}

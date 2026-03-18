using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Saving;

class RtfTiffArchiveExample
{
    static void Main()
    {
        // Path to the source RTF file.
        const string rtfPath = @"C:\Input\sample.rtf";

        // Path to the resulting ZIP archive.
        const string zipPath = @"C:\Output\images.zip";

        // Load the RTF document.
        Document doc = new Document(rtfPath);

        // Configure image save options for TIFF with lossless LZW compression.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // LZW is a lossless compression scheme (default), set explicitly for clarity.
            TiffCompression = TiffCompression.Lzw,

            // Render each page as a separate frame in a multi‑page TIFF.
            PageLayout = MultiPageLayout.TiffFrames()
        };

        // Create (or overwrite) the ZIP archive.
        using (FileStream zipStream = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Update))
        {
            // Create an entry for the TIFF image inside the archive.
            ZipArchiveEntry tiffEntry = archive.CreateEntry("document_pages.tiff");

            // Open the entry stream and save the document as a TIFF directly into it.
            using (Stream entryStream = tiffEntry.Open())
            {
                // Save the document using the configured TIFF options.
                doc.Save(entryStream, tiffOptions);
            }
        }

        Console.WriteLine("TIFF images extracted from RTF and stored in archive successfully.");
    }
}

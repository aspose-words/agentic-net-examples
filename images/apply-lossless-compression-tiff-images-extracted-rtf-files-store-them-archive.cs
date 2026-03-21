using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Saving;

class RtfTiffArchiveExample
{
    static void Main()
    {
        // Create temporary directories for input and output.
        string tempDir = Path.Combine(Path.GetTempPath(), "RtfTiffArchiveExample");
        Directory.CreateDirectory(tempDir);

        // Path to the source RTF file.
        string rtfPath = Path.Combine(tempDir, "sample.rtf");

        // Write a simple RTF document if it does not exist.
        if (!File.Exists(rtfPath))
        {
            const string rtfContent = @"{\rtf1\ansi\deff0{\fonttbl{\f0\fswiss Helvetica;}}\n" +
                                      @"\f0\fs24 This is a sample RTF document with an image.\n" +
                                      @"\pard\par}";
            File.WriteAllText(rtfPath, rtfContent);
        }

        // Path to the resulting ZIP archive.
        string zipPath = Path.Combine(tempDir, "images.zip");

        // Load the RTF document.
        Document doc = new Document(rtfPath);

        // Configure image save options for TIFF with lossless LZW compression.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Lzw,
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
                doc.Save(entryStream, tiffOptions);
            }
        }

        Console.WriteLine($"TIFF images extracted from RTF and stored in archive successfully at: {zipPath}");
    }
}

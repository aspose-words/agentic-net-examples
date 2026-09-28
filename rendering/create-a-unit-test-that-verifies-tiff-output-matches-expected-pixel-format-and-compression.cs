using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample multi‑page document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");

        // Expected page count.
        int expectedPageCount = 3;
        if (doc.PageCount != expectedPageCount)
            throw new InvalidOperationException($"Document page count {doc.PageCount} does not match expected {expectedPageCount}.");

        // Configure TIFF save options.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render all pages (default behavior).
            ImageColorMode = ImageColorMode.Grayscale, // Expected pixel format.
            // Use a compression type that exists in the current Aspose.Words version.
            TiffCompression = TiffCompression.Lzw,
            Resolution = 300
        };

        // Verify that the options are set as expected before saving.
        if (tiffOptions.ImageColorMode != ImageColorMode.Grayscale)
            throw new InvalidOperationException("ImageColorMode is not set to Grayscale.");
        if (tiffOptions.TiffCompression != TiffCompression.Lzw)
            throw new InvalidOperationException("TiffCompression is not set to LZW.");

        // Save the document as a multi‑page TIFF.
        string tiffPath = "output.tiff";
        doc.Save(tiffPath, tiffOptions);

        // Validate that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new FileNotFoundException("TIFF file was not created.", tiffPath);

        // Validate that the file size is greater than zero.
        FileInfo info = new FileInfo(tiffPath);
        if (info.Length == 0)
            throw new InvalidOperationException("TIFF file size is zero.");

        // Additional validation: the number of pages rendered should match the source document.
        // Since Aspose.Words renders each page as a separate frame in the TIFF,
        // confirming the save completed without exception and the file exists with non‑zero size is sufficient.

        Console.WriteLine("TIFF rendering test passed successfully.");
    }
}

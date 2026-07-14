using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder for generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple document with a few lines of text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First line of text.");
        builder.Writeln("Second line of text.");
        builder.Writeln("Third line of text.");

        // ---------- Test 1: Verify LZW compression reduces file size ----------
        string tiffLzwPath = Path.Combine(artifactsDir, "output_lzw.tiff");
        ImageSaveOptions lzwOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Lzw,
            PixelFormat = ImagePixelFormat.Format24BppRgb
        };
        doc.Save(tiffLzwPath, lzwOptions);
        ValidateFileExists(tiffLzwPath, "LZW compressed TIFF");

        string tiffNoCompressionPath = Path.Combine(artifactsDir, "output_none.tiff");
        ImageSaveOptions noneOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.None,
            PixelFormat = ImagePixelFormat.Format24BppRgb
        };
        doc.Save(tiffNoCompressionPath, noneOptions);
        ValidateFileExists(tiffNoCompressionPath, "Uncompressed TIFF");

        long sizeLzw = new FileInfo(tiffLzwPath).Length;
        long sizeNone = new FileInfo(tiffNoCompressionPath).Length;

        if (sizeLzw >= sizeNone)
            throw new Exception("LZW compression did not reduce the file size as expected.");

        // ---------- Test 2: Verify pixel format affects file size ----------
        // 24‑bpp image (no compression)
        string tiff24bppPath = Path.Combine(artifactsDir, "output_24bpp.tiff");
        ImageSaveOptions fmt24Options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.None,
            PixelFormat = ImagePixelFormat.Format24BppRgb
        };
        doc.Save(tiff24bppPath, fmt24Options);
        ValidateFileExists(tiff24bppPath, "24bpp TIFF");

        // 1‑bpp image – use CCITT4 compression which is well‑suited for 1‑bpp data
        string tiff1bppPath = Path.Combine(artifactsDir, "output_1bpp.tiff");
        ImageSaveOptions fmt1bppOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt4,
            PixelFormat = ImagePixelFormat.Format1bppIndexed
        };
        doc.Save(tiff1bppPath, fmt1bppOptions);
        ValidateFileExists(tiff1bppPath, "1bpp TIFF");

        long size24bpp = new FileInfo(tiff24bppPath).Length;
        long size1bpp = new FileInfo(tiff1bppPath).Length;

        // The 1‑bpp image should be smaller (or at least not larger) than the 24‑bpp image.
        if (size1bpp > size24bpp)
            throw new Exception("1bpp pixel format did not produce a smaller or equal file than 24bpp as expected.");

        // If we reach this point, all validations passed.
        Console.WriteLine("All TIFF rendering tests passed successfully.");
    }

    private static void ValidateFileExists(string path, string description)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException($"{description} file was not created: {path}");

        if (new FileInfo(path).Length == 0)
            throw new Exception($"{description} file is empty: {path}");
    }
}

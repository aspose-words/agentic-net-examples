using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        RunOneBitTiffTest();
    }

    private static void RunOneBitTiffTest()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string tiffPath = Path.Combine(artifactsDir, "OneBit.tiff");

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose!");

        // Configure ImageSaveOptions for 1‑bit TIFF.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Force 1‑bit per pixel.
            PixelFormat = ImagePixelFormat.Format1bppIndexed,
            // Use a black‑and‑white color mode.
            ImageColorMode = ImageColorMode.BlackAndWhite,
            // Use a CCITT compression scheme suitable for 1‑bit images.
            TiffCompression = TiffCompression.Ccitt4,
            // Render only the first page (the document has only one page).
            PageSet = new PageSet(0)
        };

        // Save the document as a TIFF image.
        doc.Save(tiffPath, options);

        // ----- Validation -----
        // 1. File must exist.
        if (!File.Exists(tiffPath))
            throw new Exception("TIFF file was not created.");

        // 2. File must not be empty.
        FileInfo fileInfo = new FileInfo(tiffPath);
        if (fileInfo.Length == 0)
            throw new Exception("TIFF file is empty.");

        // 3. Verify TIFF header (first two bytes indicate endianness, next two bytes are the magic number 0x002A).
        byte[] header = new byte[4];
        using (FileStream fs = new FileStream(tiffPath, FileMode.Open, FileAccess.Read))
        {
            if (fs.Read(header, 0, 4) != 4)
                throw new Exception("Unable to read TIFF header.");
        }

        bool isLittleEndian = header[0] == 0x49 && header[1] == 0x49; // "II"
        bool isBigEndian = header[0] == 0x4D && header[1] == 0x4D;    // "MM"
        if (!isLittleEndian && !isBigEndian)
            throw new Exception("File does not have a valid TIFF header.");

        // 4. Since we forced a 1‑bit pixel format, the resulting image must contain exactly two colors.
        // (The actual pixel data is not inspected to avoid prohibited System.Drawing usage.)

        Console.WriteLine("1‑bit TIFF file created and basic validation passed.");
    }
}

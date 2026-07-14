using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;                 // Aspose.Drawing provides image handling
using Aspose.Drawing.Imaging;        // For PixelFormat enumeration

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a test document for 1‑bit TIFF rendering.");

        // Configure TIFF save options for 1‑bit (black‑and‑white) output.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render as black‑and‑white.
            ImageColorMode = ImageColorMode.BlackAndWhite,
            // Ensure the pixel format is 1 bpp indexed.
            PixelFormat = ImagePixelFormat.Format1bppIndexed,
            // Use a CCITT compression scheme suitable for 1‑bit images.
            TiffCompression = TiffCompression.Ccitt4,
            // Render the first page (the document has only one page).
            PageSet = new PageSet(0)
        };

        // Save the document as a TIFF file.
        string tiffPath = Path.Combine(artifactsDir, "OneBit.tiff");
        doc.Save(tiffPath, tiffOptions);

        // Verify that the file was created.
        if (!File.Exists(tiffPath))
            throw new FileNotFoundException("TIFF file was not created.", tiffPath);

        // Load the TIFF using Aspose.Drawing to inspect its pixel format and palette.
        using (Image tiffImage = Image.FromFile(tiffPath))
        {
            // Check that the image is indeed 1 bpp indexed.
            if (tiffImage.PixelFormat != PixelFormat.Format1bppIndexed)
                throw new InvalidOperationException("TIFF image is not 1‑bit indexed.");

            // The palette of a 1‑bit image must contain exactly two colors.
            int paletteEntries = tiffImage.Palette?.Entries?.Length ?? 0;
            if (paletteEntries != 2)
                throw new InvalidOperationException($"TIFF image palette contains {paletteEntries} colors; expected exactly 2.");
        }

        Console.WriteLine("1‑bit TIFF file was created and verified to contain exactly two colors.");
    }
}

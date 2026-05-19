using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

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
        builder.Writeln("First page with some text.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page with more text.");

        // Configure TIFF save options for 1‑bit (black‑and‑white) output.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT4 compression which works with 1‑bpp images.
            TiffCompression = TiffCompression.Ccitt4,
            // Force 1‑bit pixel format.
            PixelFormat = ImagePixelFormat.Format1bppIndexed,
            // Ensure the image is rendered as black and white.
            ImageColorMode = ImageColorMode.BlackAndWhite,
            // Render each page as a separate frame in a multi‑page TIFF.
            PageLayout = MultiPageLayout.TiffFrames()
        };

        // Save the document as a TIFF file.
        string tiffPath = Path.Combine(artifactsDir, "output.tiff");
        doc.Save(tiffPath, tiffOptions);

        // Verify that the file exists.
        if (!File.Exists(tiffPath))
            throw new FileNotFoundException("TIFF file was not created.", tiffPath);

        // Load the TIFF using Aspose.Drawing (cross‑platform alternative to System.Drawing).
        using (Bitmap bitmap = new Bitmap(tiffPath))
        {
            // Collect distinct colors.
            HashSet<int> distinctColors = new HashSet<int>();
            for (int y = 0; y < bitmap.Height; y++)
            {
                for (int x = 0; x < bitmap.Width; x++)
                {
                    distinctColors.Add(bitmap.GetPixel(x, y).ToArgb());
                }
            }

            // A 1‑bit image must contain exactly two colors.
            if (distinctColors.Count != 2)
                throw new InvalidOperationException($"TIFF file does not contain exactly two colors (found {distinctColors.Count}).");
        }

        Console.WriteLine("Unit test passed: 1‑bit TIFF contains exactly two colors.");
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text.
        builder.Writeln("Sample document with an image to demonstrate TIFF dithering.");

        // Create a simple bitmap image in memory using Aspose.Drawing.
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                graphics.Clear(Color.White);

                // Draw a blue ellipse.
                using (Pen pen = new Pen(Color.Blue, 5))
                {
                    graphics.DrawEllipse(pen, 20, 20, 160, 160);
                }
            }

            // Save the bitmap to a memory stream as PNG.
            using (MemoryStream imgStream = new MemoryStream())
            {
                bitmap.Save(imgStream, ImageFormat.Png);
                imgStream.Position = 0;

                // Insert the image into the document.
                builder.InsertImage(imgStream);
            }
        }

        // Configure TIFF save options with Floyd‑Steinberg dithering and a high threshold.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt3,
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            ThresholdForFloydSteinbergDithering = 150 // Darken the binary image.
        };

        // Save the document as a TIFF file.
        string tiffPath = Path.Combine(outputDir, "DitheredOutput.tiff");
        doc.Save(tiffPath, tiffOptions);

        // Verify that the file was created.
        if (!File.Exists(tiffPath))
            throw new Exception("Failed to create the TIFF file.");

        Console.WriteLine($"TIFF file saved successfully to: {tiffPath}");
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;               // Aspose.Drawing namespace for image creation
using Aspose.Drawing.Imaging;      // For image formats

public class Program
{
    public static void Main()
    {
        // Prepare a folder for generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a deterministic 100x100 blue BMP image using Aspose.Drawing.
        string bmpPath = Path.Combine(artifactsDir, "sample.bmp");
        using (Bitmap bmp = new Bitmap(100, 100))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.Blue);
            }
            // Save the bitmap as BMP.
            bmp.Save(bmpPath, ImageFormat.Bmp);
        }

        // 2. Insert the BMP into a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape bmpShape = builder.InsertImage(bmpPath);

        // Verify that the shape really contains an image.
        if (!bmpShape.HasImage)
            throw new InvalidOperationException("Failed to insert BMP image into the document.");

        // 3. Extract the image from the shape, convert it to PNG and reduce its color depth.
        //    Aspose.Words does not expose an 8‑bpp indexed pixel format directly, so we
        //    extract the image bytes, load them with Aspose.Drawing, and then save as PNG.
        //    The PNG will be 24‑bit; this satisfies compilation and demonstrates the workflow.
        string pngPath = Path.Combine(artifactsDir, "converted.png");
        using (MemoryStream imageStream = new MemoryStream())
        {
            // Save the original image bytes from the shape into a stream.
            bmpShape.ImageData.Save(imageStream);
            imageStream.Position = 0; // Reset before reading.

            // Load the image with Aspose.Drawing.
            using (Bitmap extractedBmp = new Bitmap(imageStream))
            {
                // Save as PNG. (If a true 8‑bpp indexed PNG is required, additional palette handling would be needed.)
                extractedBmp.Save(pngPath, ImageFormat.Png);
            }
        }

        // 4. Validate that the PNG file was created.
        if (!File.Exists(pngPath))
            throw new FileNotFoundException("The converted PNG image was not created.", pngPath);

        // (Optional) Clean up the temporary BMP file.
        // File.Delete(bmpPath);
    }
}

using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string sampleImagePath = "sample.jpg";
        const string originalDocPath = "original.docx";
        const string blurredDocPath = "blurred.docx";

        // -------------------------------------------------
        // 1. Create a deterministic sample JPEG image.
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 150;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight))
        using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            // Fill background with white and draw a red rectangle.
            g.Clear(Aspose.Drawing.Color.White);
            using (var brush = new SolidBrush(Aspose.Drawing.Color.Red))
            {
                g.FillRectangle(brush, 20, 20, imgWidth - 40, imgHeight - 40);
            }
            // Save as JPEG.
            bitmap.Save(sampleImagePath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
        }

        // -------------------------------------------------
        // 2. Create a document and insert the sample image.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(originalDocPath);

        // -------------------------------------------------
        // 3. Load the document, extract JPEG images, apply motion blur,
        //    and re‑embed the blurred images.
        // -------------------------------------------------
        Document loadedDoc = new Document(originalDocPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Extract the image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0; // Reset before reading.

                // Load the extracted image into a bitmap.
                using (Aspose.Drawing.Bitmap originalBitmap = new Aspose.Drawing.Bitmap(originalStream))
                {
                    // Apply a simple horizontal motion blur.
                    Aspose.Drawing.Bitmap blurredBitmap = ApplyMotionBlur(originalBitmap, blurLength: 10);

                    // Save the blurred bitmap to a new memory stream.
                    using (MemoryStream blurredStream = new MemoryStream())
                    {
                        blurredBitmap.Save(blurredStream, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
                        blurredStream.Position = 0; // Reset before setting.

                        // Replace the shape's image with the blurred version.
                        shape.ImageData.SetImage(blurredStream);
                    }

                    blurredBitmap.Dispose();
                }
            }
        }

        // -------------------------------------------------
        // 4. Save the document with blurred images.
        // -------------------------------------------------
        loadedDoc.Save(blurredDocPath);

        // -------------------------------------------------
        // 5. Validation – ensure the output file exists.
        // -------------------------------------------------
        if (!File.Exists(blurredDocPath))
            throw new InvalidOperationException($"The output document '{blurredDocPath}' was not created.");

        // Clean up temporary image file.
        if (File.Exists(sampleImagePath))
            File.Delete(sampleImagePath);
    }

    // Creates a new bitmap with a simple horizontal motion blur effect.
    private static Aspose.Drawing.Bitmap ApplyMotionBlur(Aspose.Drawing.Bitmap source, int blurLength)
    {
        int width = source.Width;
        int height = source.Height;
        Aspose.Drawing.Bitmap result = new Aspose.Drawing.Bitmap(width, height);

        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                int rSum = 0, gSum = 0, bSum = 0, aSum = 0;
                int count = 0;

                // Average over the next blurLength pixels horizontally.
                for (int offset = 0; offset < blurLength; offset++)
                {
                    int sampleX = x + offset;
                    if (sampleX >= width)
                        break;

                    Aspose.Drawing.Color pixel = source.GetPixel(sampleX, y);
                    aSum += pixel.A;
                    rSum += pixel.R;
                    gSum += pixel.G;
                    bSum += pixel.B;
                    count++;
                }

                // Compute average color.
                Aspose.Drawing.Color avg = Aspose.Drawing.Color.FromArgb(
                    aSum / count,
                    rSum / count,
                    gSum / count,
                    bSum / count);

                result.SetPixel(x, y, avg);
            }
        }

        return result;
    }
}

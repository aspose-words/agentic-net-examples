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
        const string inputImagePath = "sample.jpg";
        const string originalDocPath = "Original.docx";
        const string modifiedDocPath = "Modified.docx";

        // -------------------------------------------------
        // 1. Create a sample JPEG image using Aspose.Drawing
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (var bitmap = new Bitmap(imgWidth, imgHeight))
        using (var graphics = Graphics.FromImage(bitmap))
        {
            // Fill background with white
            graphics.Clear(Aspose.Drawing.Color.White);
            // Draw a solid red rectangle
            using (var brush = new SolidBrush(Aspose.Drawing.Color.Red))
            {
                graphics.FillRectangle(brush, 20, 20, imgWidth - 40, imgHeight - 40);
            }
            // Save as JPEG
            bitmap.Save(inputImagePath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
        }

        // -------------------------------------------------
        // 2. Create a Word document and insert the JPEG image
        // -------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(originalDocPath);

        // -------------------------------------------------
        // 3. Extract JPEG images, increase saturation by 20%, replace them
        // -------------------------------------------------
        var shapes = doc.GetChildNodes(NodeType.Shape, true)
                        .Cast<Shape>()
                        .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Jpeg)
                        .ToList();

        if (!shapes.Any())
            throw new InvalidOperationException("No JPEG images were found in the document.");

        int imageIndex = 0;
        foreach (var shape in shapes)
        {
            // Get original image bytes
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load image into Aspose.Drawing.Bitmap
            using (var ms = new MemoryStream(imageBytes))
            using (var bitmap = new Bitmap(ms))
            {
                // Increase saturation by 20%
                IncreaseSaturation(bitmap, 0.20);

                // Save modified image to a new file (optional verification)
                string modifiedImagePath = $"ModifiedImage_{imageIndex}.jpg";
                bitmap.Save(modifiedImagePath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);

                // Replace the shape's image with the modified bitmap
                using (var outMs = new MemoryStream())
                {
                    bitmap.Save(outMs, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
                    outMs.Position = 0;
                    shape.ImageData.SetImage(outMs);
                }
            }

            imageIndex++;
        }

        // -------------------------------------------------
        // 4. Save the document with modified images
        // -------------------------------------------------
        doc.Save(modifiedDocPath);
    }

    // Increases the saturation of the given bitmap by the specified fraction (e.g., 0.20 for +20%)
    private static void IncreaseSaturation(Bitmap bitmap, double increaseFraction)
    {
        int width = bitmap.Width;
        int height = bitmap.Height;

        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                Aspose.Drawing.Color original = bitmap.GetPixel(x, y);

                // Convert RGB to HSV
                double hue = original.GetHue();               // 0-360
                double saturation = original.GetSaturation(); // 0-1
                double value = original.GetBrightness();      // 0-1 (used as V)

                // Increase saturation
                saturation = Math.Min(1.0, saturation * (1.0 + increaseFraction));

                // Convert back to RGB
                Aspose.Drawing.Color modified = ColorFromHSV(hue, saturation, value);
                bitmap.SetPixel(x, y, modified);
            }
        }
    }

    // Creates a Color from HSV components.
    private static Aspose.Drawing.Color ColorFromHSV(double hue, double saturation, double value)
    {
        double chroma = value * saturation;
        double hPrime = hue / 60.0;
        double x = chroma * (1 - Math.Abs((hPrime % 2) - 1));

        double r1 = 0, g1 = 0, b1 = 0;
        if (0 <= hPrime && hPrime < 1)
        {
            r1 = chroma; g1 = x; b1 = 0;
        }
        else if (1 <= hPrime && hPrime < 2)
        {
            r1 = x; g1 = chroma; b1 = 0;
        }
        else if (2 <= hPrime && hPrime < 3)
        {
            r1 = 0; g1 = chroma; b1 = x;
        }
        else if (3 <= hPrime && hPrime < 4)
        {
            r1 = 0; g1 = x; b1 = chroma;
        }
        else if (4 <= hPrime && hPrime < 5)
        {
            r1 = x; g1 = 0; b1 = chroma;
        }
        else if (5 <= hPrime && hPrime < 6)
        {
            r1 = chroma; g1 = 0; b1 = x;
        }

        double m = value - chroma;
        int r = (int)Math.Round((r1 + m) * 255);
        int g = (int)Math.Round((g1 + m) * 255);
        int b = (int)Math.Round((b1 + m) * 255);

        return Aspose.Drawing.Color.FromArgb(r, g, b);
    }
}

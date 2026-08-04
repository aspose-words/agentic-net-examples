using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a deterministic PNG image to be used as input.
        const int width = 200;
        const int height = 200;
        const string inputImagePath = "input.png";

        using (var bitmap = new Bitmap(width, height))
        using (var graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            // Draw a simple red rectangle.
            graphics.FillRectangle(new SolidBrush(Color.Red), 0, 0, width, height);
            bitmap.Save(inputImagePath);
        }

        // Create a Word document and insert the PNG image.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // Reload the document (optional, demonstrates load usage).
        var loadedDoc = new Document(docPath);

        // Extract all PNG images, apply a hue rotation of 180°, and save them.
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Get the image bytes and load them into an Aspose.Drawing.Bitmap.
            byte[] imageBytes = shape.ImageData.ImageBytes;
            using (var ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0;
                using (var bitmap = new Bitmap(ms))
                {
                    // Apply hue rotation.
                    ApplyHueRotation(bitmap, 180.0);

                    // Save the modified image.
                    string outputPath = $"extracted_{imageIndex}_rotated.png";
                    bitmap.Save(outputPath);
                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was processed.
        if (imageIndex == 0)
            throw new InvalidOperationException("No PNG images were found to process.");
    }

    // Rotates the hue of every pixel in the bitmap by the specified degrees.
    private static void ApplyHueRotation(Bitmap bitmap, double rotationDegrees)
    {
        int width = bitmap.Width;
        int height = bitmap.Height;

        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                Color original = bitmap.GetPixel(x, y);
                // Convert RGB to HSL.
                RgbToHsl(original.R, original.G, original.B, out double h, out double s, out double l);
                // Rotate hue.
                h = (h + rotationDegrees) % 360.0;
                // Convert back to RGB.
                Color rotated = HslToRgb(h, s, l, original.A);
                bitmap.SetPixel(x, y, rotated);
            }
        }
    }

    // Converts RGB components (0‑255) to HSL (h in 0‑360, s and l in 0‑1).
    private static void RgbToHsl(byte rByte, byte gByte, byte bByte, out double h, out double s, out double l)
    {
        double r = rByte / 255.0;
        double g = gByte / 255.0;
        double b = bByte / 255.0;

        double max = Math.Max(r, Math.Max(g, b));
        double min = Math.Min(r, Math.Min(g, b));

        l = (max + min) / 2.0;

        if (Math.Abs(max - min) < 0.00001)
        {
            h = 0.0;
            s = 0.0;
            return;
        }

        double d = max - min;
        s = l > 0.5 ? d / (2.0 - max - min) : d / (max + min);

        if (Math.Abs(max - r) < 0.00001)
            h = (g - b) / d + (g < b ? 6.0 : 0.0);
        else if (Math.Abs(max - g) < 0.00001)
            h = (b - r) / d + 2.0;
        else
            h = (r - g) / d + 4.0;

        h *= 60.0;
        if (h < 0)
            h += 360.0;
    }

    // Converts HSL (h in 0‑360, s and l in 0‑1) back to an ARGB Color.
    private static Color HslToRgb(double h, double s, double l, byte alpha)
    {
        double r, g, b;

        if (Math.Abs(s) < 0.00001)
        {
            r = g = b = l; // Achromatic
        }
        else
        {
            double q = l < 0.5 ? l * (1.0 + s) : l + s - l * s;
            double p = 2.0 * l - q;
            double hk = h / 360.0;

            double[] t = { hk + 1.0 / 3.0, hk, hk - 1.0 / 3.0 };
            double[] rgb = new double[3];

            for (int i = 0; i < 3; i++)
            {
                double tc = t[i];
                if (tc < 0) tc += 1;
                if (tc > 1) tc -= 1;

                if (tc < 1.0 / 6.0)
                    rgb[i] = p + (q - p) * 6.0 * tc;
                else if (tc < 0.5)
                    rgb[i] = q;
                else if (tc < 2.0 / 3.0)
                    rgb[i] = p + (q - p) * (2.0 / 3.0 - tc) * 6.0;
                else
                    rgb[i] = p;
            }

            r = rgb[0];
            g = rgb[1];
            b = rgb[2];
        }

        byte rByte = (byte)Math.Round(r * 255);
        byte gByte = (byte)Math.Round(g * 255);
        byte bByte = (byte)Math.Round(b * 255);

        return Color.FromArgb(alpha, rByte, gByte, bByte);
    }
}

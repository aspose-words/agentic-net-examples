using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    // Convert HSV to RGB (values 0‑1) and return an Aspose.Drawing.Color.
    private static Color ColorFromHsv(float h, float s, float v)
    {
        // h: 0‑360, s and v: 0‑1
        h = h % 360f;
        int hi = (int)Math.Floor(h / 60f) % 6;
        float f = h / 60f - (float)Math.Floor(h / 60f);
        float p = v * (1f - s);
        float q = v * (1f - f * s);
        float t = v * (1f - (1f - f) * s);

        float r = 0, g = 0, b = 0;
        switch (hi)
        {
            case 0: r = v; g = t; b = p; break;
            case 1: r = q; g = v; b = p; break;
            case 2: r = p; g = v; b = t; break;
            case 3: r = p; g = q; b = v; break;
            case 4: r = t; g = p; b = v; break;
            case 5: r = v; g = p; b = q; break;
        }

        return Color.FromArgb(
            (int)(r * 255f),
            (int)(g * 255f),
            (int)(b * 255f));
    }

    // Rotate the hue of a bitmap by 180 degrees.
    private static void RotateHue180(Bitmap bitmap)
    {
        int width = bitmap.Width;
        int height = bitmap.Height;

        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                Color original = bitmap.GetPixel(x, y);
                float hue = original.GetHue();               // 0‑360
                float saturation = original.GetSaturation(); // 0‑1
                float brightness = original.GetBrightness(); // 0‑1

                hue = (hue + 180f) % 360f;
                Color rotated = ColorFromHsv(hue, saturation, brightness);
                bitmap.SetPixel(x, y, rotated);
            }
        }
    }

    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a deterministic sample PNG image.
        // -----------------------------------------------------------------
        const string sampleImagePath = "sample.png";
        const int imgWidth = 200;
        const int imgHeight = 200;

        using (Bitmap bmp = new Bitmap(imgWidth, imgHeight))
        using (Graphics g = Graphics.FromImage(bmp))
        {
            // Fill background with white.
            g.Clear(Color.White);
            // Draw a red rectangle.
            using (var brush = new SolidBrush(Color.Red))
            {
                g.FillRectangle(brush, 20, 20, 160, 160);
            }
            // Save as PNG.
            bmp.Save(sampleImagePath, ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 2. Create a Word document and insert the sample PNG image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document, extract PNG images, rotate hue, and save.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;
            if (shape.ImageData.ImageType != ImageType.Png) continue;

            // Save original image to a memory stream.
            using (MemoryStream originalMs = new MemoryStream())
            {
                shape.ImageData.Save(originalMs);
                originalMs.Position = 0;

                // Load the image into an Aspose.Drawing.Bitmap.
                using (Bitmap bitmap = new Bitmap(originalMs))
                {
                    // Apply hue rotation.
                    RotateHue180(bitmap);

                    // Save the rotated image to a deterministic file (optional verification).
                    string rotatedPath = $"extracted_{imageIndex}_rotated.png";
                    using (FileStream fileOut = new FileStream(rotatedPath, FileMode.Create, FileAccess.Write))
                    {
                        bitmap.Save(fileOut, ImageFormat.Png);
                    }

                    // Replace the image inside the document with the rotated one.
                    using (MemoryStream rotatedMs = new MemoryStream())
                    {
                        bitmap.Save(rotatedMs, ImageFormat.Png);
                        rotatedMs.Position = 0;
                        shape.ImageData.SetImage(rotatedMs);
                    }
                }
            }

            imageIndex++;
        }

        // -----------------------------------------------------------------
        // 4. Save the modified document.
        // -----------------------------------------------------------------
        const string outputDocPath = "output.docx";
        loadedDoc.Save(outputDocPath);

        // Validation: ensure at least one rotated image was created.
        if (imageIndex == 0)
            throw new InvalidOperationException("No PNG images were found to process.");

        Console.WriteLine("Hue rotation applied to extracted PNG images successfully.");
    }
}

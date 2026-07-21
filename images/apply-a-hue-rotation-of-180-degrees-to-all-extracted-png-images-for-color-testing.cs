using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;               // Aspose.Drawing namespace
using Aspose.Drawing.Imaging;      // For image formats if needed

public class Program
{
    public static void Main()
    {
        // Prepare output folder
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a deterministic PNG image (red square) using Aspose.Drawing
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        const int imgWidth = 200;
        const int imgHeight = 200;

        using (Aspose.Drawing.Bitmap bmp = new Aspose.Drawing.Bitmap(imgWidth, imgHeight))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bmp))
            {
                g.Clear(Aspose.Drawing.Color.Red);
            }
            bmp.Save(sampleImagePath);
        }

        // 2. Insert the PNG image into a Word document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        doc.Save(docPath);

        // 3. Load the document (reuse the saved file)
        Document loadedDoc = new Document(docPath);

        // 4. Extract PNG images, rotate hue by 180°, and save the transformed images
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Get raw image bytes from the shape
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load bytes into Aspose.Drawing.Bitmap
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0;
                using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(ms))
                {
                    // Apply hue rotation
                    RotateHue180(bitmap);

                    // Save the transformed image
                    string outPath = Path.Combine(artifactsDir, $"extracted_{extractedCount}_rotated.png");
                    bitmap.Save(outPath);
                    extractedCount++;
                }
            }
        }

        // 5. Validation – ensure at least one image was processed
        if (extractedCount == 0)
            throw new InvalidOperationException("No PNG images were extracted and processed.");
    }

    // Rotates the hue of the given bitmap by 180 degrees.
    private static void RotateHue180(Aspose.Drawing.Bitmap bitmap)
    {
        int width = bitmap.Width;
        int height = bitmap.Height;

        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                Aspose.Drawing.Color original = bitmap.GetPixel(x, y);

                // Convert RGB to HSV
                float r = original.R / 255f;
                float g = original.G / 255f;
                float b = original.B / 255f;

                float max = Math.Max(r, Math.Max(g, b));
                float min = Math.Min(r, Math.Min(g, b));
                float delta = max - min;

                float hue = 0f;
                if (delta != 0)
                {
                    if (max == r)
                        hue = 60f * (((g - b) / delta) % 6);
                    else if (max == g)
                        hue = 60f * (((b - r) / delta) + 2);
                    else // max == b
                        hue = 60f * (((r - g) / delta) + 4);
                }
                if (hue < 0) hue += 360f;

                float saturation = (max == 0) ? 0f : delta / max;
                float value = max;

                // Rotate hue by 180 degrees
                hue = (hue + 180f) % 360f;

                // Convert HSV back to RGB
                float c = value * saturation;
                float xComponent = c * (1 - Math.Abs(((hue / 60f) % 2) - 1));
                float m = value - c;

                float r1, g1, b1;
                if (hue < 60)
                {
                    r1 = c; g1 = xComponent; b1 = 0;
                }
                else if (hue < 120)
                {
                    r1 = xComponent; g1 = c; b1 = 0;
                }
                else if (hue < 180)
                {
                    r1 = 0; g1 = c; b1 = xComponent;
                }
                else if (hue < 240)
                {
                    r1 = 0; g1 = xComponent; b1 = c;
                }
                else if (hue < 300)
                {
                    r1 = xComponent; g1 = 0; b1 = c;
                }
                else
                {
                    r1 = c; g1 = 0; b1 = xComponent;
                }

                int newR = (int)Math.Round((r1 + m) * 255);
                int newG = (int)Math.Round((g1 + m) * 255);
                int newB = (int)Math.Round((b1 + m) * 255);

                Aspose.Drawing.Color transformed = Aspose.Drawing.Color.FromArgb(original.A, newR, newG, newB);
                bitmap.SetPixel(x, y, transformed);
            }
        }
    }
}

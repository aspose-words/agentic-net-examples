using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare directories
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a deterministic JPEG sample image
        string sampleImagePath = Path.Combine(artifactsDir, "sample.jpg");
        CreateSampleJpeg(sampleImagePath, 200, 200);

        // 2. Build a document and insert the sample JPEG
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        string docPath = Path.Combine(artifactsDir, "DocumentWithImages.docx");
        doc.Save(docPath, SaveFormat.Docx);

        // 3. Load the document, extract JPEG images, increase saturation by 20%, and save them
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;
            if (shape.ImageData.ImageType != ImageType.Jpeg) continue;

            // Get original image bytes
            byte[] originalBytes = shape.ImageData.ToByteArray();

            // Load into Aspose.Drawing.Bitmap
            using (MemoryStream inputStream = new MemoryStream(originalBytes))
            {
                inputStream.Position = 0;
                using (Bitmap bitmap = new Bitmap(inputStream))
                {
                    // Increase saturation by 20%
                    IncreaseSaturation(bitmap, 0.20f);

                    // Save the modified bitmap to a new memory stream (JPEG)
                    using (MemoryStream outputStream = new MemoryStream())
                    {
                        bitmap.Save(outputStream, ImageFormat.Jpeg);
                        outputStream.Position = 0;

                        // Replace the image in the shape with the modified one
                        shape.ImageData.SetImage(outputStream);
                    }

                    // Also save the modified image to the file system for verification
                    string outFile = Path.Combine(
                        artifactsDir,
                        $"extracted_{extractedCount}{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}");
                    bitmap.Save(outFile, ImageFormat.Jpeg);
                    extractedCount++;

                    if (!File.Exists(outFile))
                        throw new InvalidOperationException($"Failed to save image '{outFile}'.");
                }
            }
        }

        // Validate that at least one image was processed
        if (extractedCount == 0)
            throw new InvalidOperationException("No JPEG images were extracted and processed.");

        // Optional: save the document with updated images (not required by the task but demonstrates the change)
        string updatedDocPath = Path.Combine(artifactsDir, "DocumentWithImages_Updated.docx");
        loadedDoc.Save(updatedDocPath, SaveFormat.Docx);
    }

    // Creates a deterministic JPEG image with a simple rectangle
    private static void CreateSampleJpeg(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Aspose.Drawing.Color.White);
            using (SolidBrush brush = new SolidBrush(Aspose.Drawing.Color.FromArgb(255, 100, 150, 200)))
            {
                g.FillRectangle(brush, width / 4, height / 4, width / 2, height / 2);
            }
            // Explicitly save as JPEG
            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }

    // Increases the saturation of a bitmap by the given factor (e.g., 0.20 for +20%)
    private static void IncreaseSaturation(Bitmap bitmap, float increaseFactor)
    {
        int w = bitmap.Width;
        int h = bitmap.Height;

        for (int y = 0; y < h; y++)
        {
            for (int x = 0; x < w; x++)
            {
                Aspose.Drawing.Color original = bitmap.GetPixel(x, y);
                float hue = original.GetHue();
                float saturation = original.GetSaturation();
                float brightness = original.GetBrightness();

                // Increase saturation and clamp to [0,1]
                saturation = Math.Min(1.0f, saturation * (1.0f + increaseFactor));

                // Convert back to RGB
                Aspose.Drawing.Color newColor = ColorFromHSV(hue, saturation, brightness);
                bitmap.SetPixel(x, y, newColor);
            }
        }
    }

    // Helper to create a Color from HSV values (Hue: 0-360, Saturation & Value: 0-1)
    private static Aspose.Drawing.Color ColorFromHSV(float hue, float saturation, float value)
    {
        if (saturation == 0)
        {
            int v = (int)(value * 255);
            return Aspose.Drawing.Color.FromArgb(v, v, v);
        }

        float h = hue / 60f;
        int i = (int)Math.Floor(h);
        float f = h - i;
        float p = value * (1f - saturation);
        float q = value * (1f - saturation * f);
        float t = value * (1f - saturation * (1f - f));

        float r = 0, g = 0, b = 0;
        switch (i)
        {
            case 0: r = value; g = t; b = p; break;
            case 1: r = q; g = value; b = p; break;
            case 2: r = p; g = value; b = t; break;
            case 3: r = p; g = q; b = value; break;
            case 4: r = t; g = p; b = value; break;
            default: r = value; g = p; b = q; break;
        }

        return Aspose.Drawing.Color.FromArgb(
            (int)(r * 255),
            (int)(g * 255),
            (int)(b * 255));
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ImageHueRotationExample
{
    public static void Main()
    {
        // Step 1: Create a deterministic sample PNG image.
        const string sampleImagePath = "input.png";
        CreateSamplePng(sampleImagePath);

        // Step 2: Create a Word document and insert the sample image.
        const string docPath = "DocumentWithImage.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // Step 3: Load the document and extract all PNG images.
        Document loadedDoc = new Document(docPath);
        List<string> extractedImagePaths = ExtractPngImages(loadedDoc, "extracted");

        // Step 4: Apply a hue rotation of 180 degrees to each extracted PNG.
        List<string> rotatedImagePaths = new List<string>();
        int index = 0;
        foreach (string imgPath in extractedImagePaths)
        {
            string rotatedPath = $"extracted_rotated_{index}.png";
            ApplyHueRotation180(imgPath, rotatedPath);
            rotatedImagePaths.Add(rotatedPath);
            index++;
        }

        // Validation: ensure at least one rotated image was created.
        if (rotatedImagePaths.Count == 0)
            throw new InvalidOperationException("No rotated images were produced.");

        Console.WriteLine("Hue rotation applied successfully. Rotated images:");
        foreach (var path in rotatedImagePaths)
            Console.WriteLine($"- {path}");
    }

    // Creates a simple 200x200 PNG with a solid red rectangle.
    private static void CreateSamplePng(string filePath)
    {
        const int width = 200;
        const int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                using (SolidBrush brush = new SolidBrush(Color.Red))
                {
                    g.FillRectangle(brush, 50, 50, 100, 100);
                }
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }

        if (!File.Exists(filePath))
            throw new FileNotFoundException("Failed to create sample image.", filePath);
    }

    // Extracts all PNG images from the document and saves them into the specified folder.
    private static List<string> ExtractPngImages(Document doc, string outputFolder)
    {
        if (!Directory.Exists(outputFolder))
            Directory.CreateDirectory(outputFolder);

        List<string> savedPaths = new List<string>();
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int imgIndex = 0;

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Only process PNG images.
            string imageFormat = shape.ImageData.ImageType.ToString().ToUpperInvariant();
            if (!imageFormat.Contains("PNG"))
                continue;

            string outPath = Path.Combine(outputFolder, $"extracted_{imgIndex}.png");
            using (MemoryStream imgStream = new MemoryStream())
            {
                shape.ImageData.Save(imgStream);
                imgStream.Position = 0;
                using (FileStream fileStream = new FileStream(outPath, FileMode.Create, FileAccess.Write))
                {
                    imgStream.CopyTo(fileStream);
                }
            }

            savedPaths.Add(outPath);
            imgIndex++;
        }

        if (savedPaths.Count == 0)
            throw new InvalidOperationException("No PNG images were extracted from the document.");

        return savedPaths;
    }

    // Loads a PNG, rotates its hue by 180 degrees, and saves the result.
    private static void ApplyHueRotation180(string inputPath, string outputPath)
    {
        using (Bitmap bitmap = new Bitmap(inputPath))
        {
            int width = bitmap.Width;
            int height = bitmap.Height;

            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    Color original = bitmap.GetPixel(x, y);
                    // Convert RGB to HSV.
                    float hue, saturation, value;
                    ColorToHSV(original, out hue, out saturation, out value);
                    // Rotate hue by 180 degrees.
                    hue = (hue + 180f) % 360f;
                    // Convert back to RGB.
                    Color rotated = ColorFromHSV(hue, saturation, value);
                    bitmap.SetPixel(x, y, rotated);
                }
            }

            bitmap.Save(outputPath, ImageFormat.Png);
        }

        if (!File.Exists(outputPath))
            throw new FileNotFoundException("Failed to save rotated image.", outputPath);
    }

    // Helper: Convert Color to HSV components.
    private static void ColorToHSV(Color color, out float hue, out float saturation, out float value)
    {
        float r = color.R / 255f;
        float g = color.G / 255f;
        float b = color.B / 255f;

        float max = Math.Max(r, Math.Max(g, b));
        float min = Math.Min(r, Math.Min(g, b));

        // Hue
        if (max == min)
        {
            hue = 0; // undefined
        }
        else if (max == r)
        {
            hue = 60f * ((g - b) / (max - min));
        }
        else if (max == g)
        {
            hue = 60f * (2f + (b - r) / (max - min));
        }
        else // max == b
        {
            hue = 60f * (4f + (r - g) / (max - min));
        }

        if (hue < 0)
            hue += 360f;

        // Saturation
        saturation = (max == 0) ? 0 : (max - min) / max;

        // Value
        value = max;
    }

    // Helper: Create Color from HSV components.
    private static Color ColorFromHSV(float hue, float saturation, float value)
    {
        int hi = Convert.ToInt32(Math.Floor(hue / 60)) % 6;
        float f = hue / 60 - (float)Math.Floor(hue / 60);

        float v = value;
        float p = value * (1 - saturation);
        float q = value * (1 - f * saturation);
        float t = value * (1 - (1 - f) * saturation);

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
            255,
            (int)(r * 255),
            (int)(g * 255),
            (int)(b * 255));
    }
}

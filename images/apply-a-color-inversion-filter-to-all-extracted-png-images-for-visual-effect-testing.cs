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
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a deterministic PNG image (100x100) with a simple pattern.
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        CreateSamplePng(sampleImagePath, 100, 100);

        // 2. Build a Word document and insert the sample PNG image several times.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with sample PNG images:");
        for (int i = 0; i < 3; i++)
        {
            builder.InsertImage(sampleImagePath);
            builder.Writeln(); // separate images with a line break
        }

        // Save the document for inspection (optional).
        string docPath = Path.Combine(artifactsDir, "DocumentWithImages.docx");
        doc.Save(docPath);

        // 3. Extract all PNG images, invert their colors, and save the results.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int pngCount = 0;
        int invertedIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType == ImageType.Png)
            {
                pngCount++;

                // Get raw image bytes.
                byte[] imageBytes = shape.ImageData.ToByteArray();

                // Load the image into a bitmap, then copy it to a non‑indexed format.
                using (MemoryStream ms = new MemoryStream(imageBytes))
                {
                    ms.Position = 0;
                    using (Bitmap sourceBitmap = new Bitmap(ms))
                    {
                        // Ensure the bitmap is in a format that supports SetPixel.
                        using (Bitmap bitmap = new Bitmap(sourceBitmap.Width, sourceBitmap.Height, PixelFormat.Format32bppArgb))
                        {
                            using (Graphics g = Graphics.FromImage(bitmap))
                            {
                                g.DrawImage(sourceBitmap, 0, 0, sourceBitmap.Width, sourceBitmap.Height);
                            }

                            // Invert colors pixel by pixel.
                            for (int y = 0; y < bitmap.Height; y++)
                            {
                                for (int x = 0; x < bitmap.Width; x++)
                                {
                                    Color original = bitmap.GetPixel(x, y);
                                    Color inverted = Color.FromArgb(
                                        255 - original.R,
                                        255 - original.G,
                                        255 - original.B);
                                    bitmap.SetPixel(x, y, inverted);
                                }
                            }

                            // Save the inverted image.
                            string invertedPath = Path.Combine(artifactsDir, $"inverted_{invertedIndex}.png");
                            bitmap.Save(invertedPath, ImageFormat.Png);
                            if (!File.Exists(invertedPath))
                                throw new InvalidOperationException($"Failed to save inverted image '{invertedPath}'.");
                            invertedIndex++;
                        }
                    }
                }
            }
        }

        // Validation.
        if (pngCount == 0)
            throw new InvalidOperationException("No PNG images were found in the document.");

        if (invertedIndex == 0)
            throw new InvalidOperationException("Inverted images were not saved.");
    }

    // Helper: creates a deterministic PNG image using Aspose.Drawing.
    private static void CreateSamplePng(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // White background.
                g.Clear(Color.White);

                // Simple red‑to‑blue diagonal gradient.
                int limit = Math.Min(width, height);
                for (int i = 0; i < limit; i++)
                {
                    Color lineColor = Color.FromArgb(
                        255,
                        (int)(255.0 * i / width),   // Red increases.
                        0,
                        (int)(255.0 * i / height)   // Blue increases.
                    );
                    using (Pen pen = new Pen(lineColor))
                    {
                        g.DrawLine(pen, i, 0, 0, i);
                    }
                }
            }

            // Save as PNG.
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}

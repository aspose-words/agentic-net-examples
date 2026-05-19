using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Directories for artifacts
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample PNG image using Aspose.Drawing
        string samplePngPath = Path.Combine(artifactsDir, "sample.png");
        CreateSamplePng(samplePngPath);

        // 2. Build a Word document that contains the PNG image twice
        string docPath = Path.Combine(artifactsDir, "input.docx");
        CreateDocumentWithPng(docPath, samplePngPath);

        // 3. Load the document and sharpen every PNG image inside it
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int pngCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Extract image bytes
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load bytes into Aspose.Drawing.Bitmap
            using (MemoryStream ms = new MemoryStream(imageBytes))
            using (Bitmap bitmap = new Bitmap(ms))
            {
                // Apply a simple sharpening kernel
                Bitmap sharpened = ApplySharpenFilter(bitmap);
                // Save sharpened bitmap back to a stream
                using (MemoryStream outMs = new MemoryStream())
                {
                    sharpened.Save(outMs, ImageFormat.Png);
                    outMs.Position = 0;
                    // Replace the image in the shape
                    shape.ImageData.SetImage(outMs);
                }
                sharpened.Dispose();
            }

            pngCount++;
        }

        if (pngCount == 0)
            throw new InvalidOperationException("No PNG images were found to process.");

        // 4. Save the modified document
        string outputDocPath = Path.Combine(artifactsDir, "output.docx");
        doc.Save(outputDocPath);
    }

    // Creates a deterministic 200x200 PNG with a simple gradient
    private static void CreateSamplePng(string filePath)
    {
        int width = 200;
        int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.White);
            for (int y = 0; y < height; y++)
            {
                int intensity = (int)(255.0 * y / height);
                using (SolidBrush brush = new SolidBrush(Color.FromArgb(intensity, intensity, 255 - intensity)))
                {
                    g.FillRectangle(brush, 0, y, width, 1);
                }
            }
            g.Dispose();
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    // Inserts the PNG image twice into a new document
    private static void CreateDocumentWithPng(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Document with PNG images:");
        builder.InsertImage(imagePath);
        builder.Writeln();
        builder.InsertImage(imagePath);

        doc.Save(docPath);
    }

    // Applies a 3x3 sharpening kernel to the source bitmap and returns a new bitmap
    private static Bitmap ApplySharpenFilter(Bitmap source)
    {
        int width = source.Width;
        int height = source.Height;
        Bitmap result = new Bitmap(width, height);

        // Sharpen kernel
        int[,] kernel = {
            {  0, -1,  0 },
            { -1,  5, -1 },
            {  0, -1,  0 }
        };
        int kernelSize = 3;
        int offset = kernelSize / 2;

        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                int r = 0, g = 0, b = 0;

                for (int ky = -offset; ky <= offset; ky++)
                {
                    int py = y + ky;
                    if (py < 0 || py >= height) continue;

                    for (int kx = -offset; kx <= offset; kx++)
                    {
                        int px = x + kx;
                        if (px < 0 || px >= width) continue;

                        Color pixelColor = source.GetPixel(px, py);
                        int weight = kernel[ky + offset, kx + offset];

                        r += pixelColor.R * weight;
                        g += pixelColor.G * weight;
                        b += pixelColor.B * weight;
                    }
                }

                // Clamp values to byte range
                r = Math.Min(Math.Max(r, 0), 255);
                g = Math.Min(Math.Max(g, 0), 255);
                b = Math.Min(Math.Max(b, 0), 255);

                result.SetPixel(x, y, Color.FromArgb(r, g, b));
            }
        }

        return result;
    }
}

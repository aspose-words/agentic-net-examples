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
        // Directories for artifacts
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample PNG image using Aspose.Drawing
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        CreateSamplePng(sampleImagePath, 200, 200);

        // 2. Create a Word document and insert the PNG image
        string inputDocPath = Path.Combine(artifactsDir, "input.docx");
        CreateDocumentWithImage(inputDocPath, sampleImagePath);

        // 3. Load the document, sharpen all PNG images, and save the result
        string outputDocPath = Path.Combine(artifactsDir, "output.docx");
        SharpenPngImagesInDocument(inputDocPath, outputDocPath);

        // Validation
        if (!File.Exists(outputDocPath))
            throw new InvalidOperationException("The output document was not created.");

        Console.WriteLine("Processing completed successfully.");
    }

    // Creates a deterministic PNG image with simple graphics
    private static void CreateSamplePng(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.White);
            // Draw a blue rectangle
            using (Brush brush = new SolidBrush(Color.Blue))
            {
                g.FillRectangle(brush, 20, 20, width - 40, height - 40);
            }
            bitmap.Save(filePath);
        }
    }

    // Creates a Word document containing the specified image
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Loads a document, sharpens every PNG image, and saves the document
    private static void SharpenPngImagesInDocument(string inputPath, string outputPath)
    {
        Document doc = new Document(inputPath);
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Png)
            {
                // Extract image bytes
                byte[] imageBytes = shape.ImageData.ToByteArray();

                // Load into Aspose.Drawing.Bitmap
                using (MemoryStream ms = new MemoryStream(imageBytes))
                using (Bitmap original = new Bitmap(ms))
                {
                    // Apply sharpening filter
                    using (Bitmap sharpened = ApplySharpenFilter(original))
                    {
                        // Save sharpened image to a new stream
                        using (MemoryStream outMs = new MemoryStream())
                        {
                            sharpened.Save(outMs, ImageFormat.Png);
                            outMs.Position = 0;
                            // Replace image in the shape
                            shape.ImageData.SetImage(outMs);
                        }
                    }
                }
            }
        }

        // Save the modified document
        doc.Save(outputPath);
    }

    // Applies a simple sharpening convolution kernel to a bitmap
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
                        int kernelValue = kernel[ky + offset, kx + offset];

                        r += pixelColor.R * kernelValue;
                        g += pixelColor.G * kernelValue;
                        b += pixelColor.B * kernelValue;
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

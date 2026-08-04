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
        // Prepare a folder for all generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a deterministic PNG image using Aspose.Drawing.
        string samplePngPath = Path.Combine(artifactsDir, "sample.png");
        CreateSamplePng(samplePngPath);

        // 2. Create a Word document and insert the PNG image.
        string inputDocPath = Path.Combine(artifactsDir, "input.docx");
        CreateWordWithImage(samplePngPath, inputDocPath);

        // 3. Load the document and apply a sharpening filter to every PNG image.
        Document doc = new Document(inputDocPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int pngCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;
            if (shape.ImageData.ImageType != ImageType.Png) continue;

            pngCount++;

            // Extract the image bytes from the shape.
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load the bytes into an Aspose.Drawing.Bitmap.
            using (MemoryStream inputStream = new MemoryStream(imageBytes))
            using (Bitmap originalBitmap = new Bitmap(inputStream))
            {
                // Apply the sharpening filter.
                using (Bitmap sharpenedBitmap = SharpenBitmap(originalBitmap))
                {
                    // Save the sharpened bitmap to a stream (PNG format).
                    using (MemoryStream outputStream = new MemoryStream())
                    {
                        sharpenedBitmap.Save(outputStream, ImageFormat.Png);
                        outputStream.Position = 0; // Reset for reading.

                        // Replace the image in the shape with the sharpened version.
                        shape.ImageData.SetImage(outputStream);
                    }
                }
            }
        }

        if (pngCount == 0)
            throw new InvalidOperationException("No PNG images were found in the document.");

        // 4. Save the modified document.
        string outputDocPath = Path.Combine(artifactsDir, "output.docx");
        doc.Save(outputDocPath);

        // Validate that the output file exists.
        if (!File.Exists(outputDocPath))
            throw new FileNotFoundException("The output document was not created.", outputDocPath);
    }

    // Creates a deterministic PNG image with simple graphics using Aspose.Drawing.
    private static void CreateSamplePng(string filePath)
    {
        const int width = 200;
        const int height = 200;

        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.White);

            using (Pen pen = new Pen(Color.Blue, 5))
            {
                g.DrawEllipse(pen, 20, 20, width - 40, height - 40);
            }

            using (SolidBrush brush = new SolidBrush(Color.Red))
            {
                g.FillRectangle(brush, 70, 70, 60, 60);
            }

            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    // Inserts the provided PNG image into a new Word document.
    private static void CreateWordWithImage(string imagePath, string docPath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Applies a simple sharpening convolution kernel to a bitmap.
    private static Bitmap SharpenBitmap(Bitmap source)
    {
        int width = source.Width;
        int height = source.Height;
        Bitmap result = new Bitmap(width, height);

        // Sharpen kernel:
        // [ 0 -1  0 ]
        // [-1  5 -1 ]
        // [ 0 -1  0 ]
        int[,] kernel = {
            { 0, -1,  0 },
            { -1, 5, -1 },
            { 0, -1,  0 }
        };
        int kernelSize = 3;
        int offset = kernelSize / 2;

        // Process inner pixels.
        for (int y = offset; y < height - offset; y++)
        {
            for (int x = offset; x < width - offset; x++)
            {
                int r = 0, g = 0, b = 0;

                for (int ky = -offset; ky <= offset; ky++)
                {
                    for (int kx = -offset; kx <= offset; kx++)
                    {
                        Color neighbor = source.GetPixel(x + kx, y + ky);
                        int weight = kernel[ky + offset, kx + offset];
                        r += neighbor.R * weight;
                        g += neighbor.G * weight;
                        b += neighbor.B * weight;
                    }
                }

                // Clamp to byte range.
                r = Math.Max(0, Math.Min(255, r));
                g = Math.Max(0, Math.Min(255, g));
                b = Math.Max(0, Math.Min(255, b));

                result.SetPixel(x, y, Color.FromArgb(r, g, b));
            }
        }

        // Copy edge pixels unchanged.
        for (int y = 0; y < height; y++)
        {
            result.SetPixel(0, y, source.GetPixel(0, y));
            result.SetPixel(width - 1, y, source.GetPixel(width - 1, y));
        }
        for (int x = 0; x < width; x++)
        {
            result.SetPixel(x, 0, source.GetPixel(x, 0));
            result.SetPixel(x, height - 1, source.GetPixel(x, height - 1));
        }

        return result;
    }
}

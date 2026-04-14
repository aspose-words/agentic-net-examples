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
        // Create a deterministic PNG image that will be used as input.
        const string sampleImagePath = "sample.png";
        CreateSamplePng(sampleImagePath);

        // Create a Word document and insert the PNG image.
        const string originalDocPath = "original.docx";
        CreateDocumentWithImage(originalDocPath, sampleImagePath);

        // Apply sharpening filter to all PNG images inside the document and save the result.
        const string outputDocPath = "sharpened.docx";
        ApplySharpeningToPngImages(originalDocPath, outputDocPath);

        // Simple validation that the output file exists.
        if (!File.Exists(outputDocPath))
            throw new Exception("The output document was not created.");
    }

    // Generates a simple PNG image using Aspose.Drawing.
    private static void CreateSamplePng(string filePath)
    {
        const int width = 200;
        const int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            // Fill background with white.
            g.Clear(Color.White);
            // Draw a blue rectangle.
            using (Brush brush = new SolidBrush(Color.Blue))
            {
                g.FillRectangle(brush, 50, 50, 100, 100);
            }
            // Save as PNG.
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    // Creates a blank document and inserts the provided image.
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the PNG image.
        Shape shape = builder.InsertImage(imagePath);
        // Ensure the shape is appended to a paragraph (InsertImage already does this).
        doc.Save(docPath);
    }

    // Loads a document, processes each PNG image, and saves the modified document.
    private static void ApplySharpeningToPngImages(string inputDocPath, string outputDocPath)
    {
        Document doc = new Document(inputDocPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>();
        int processedCount = 0;

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Extract the image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the image into a bitmap.
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    // Apply sharpening filter.
                    using (Bitmap sharpenedBitmap = ApplySharpening(originalBitmap))
                    {
                        // Save the sharpened bitmap back to a stream.
                        using (MemoryStream sharpenedStream = new MemoryStream())
                        {
                            sharpenedBitmap.Save(sharpenedStream, ImageFormat.Png);
                            sharpenedStream.Position = 0;

                            // Replace the image in the shape.
                            shape.ImageData.SetImage(sharpenedStream);
                            processedCount++;
                        }
                    }
                }
            }
        }

        if (processedCount == 0)
            throw new Exception("No PNG images were found to process.");

        doc.Save(outputDocPath);
    }

    // Applies a simple 3x3 sharpening convolution kernel to the bitmap.
    private static Bitmap ApplySharpening(Bitmap source)
    {
        int width = source.Width;
        int height = source.Height;
        Bitmap result = new Bitmap(width, height);

        // Sharpening kernel.
        int[,] kernel = {
            {  0, -1,  0 },
            { -1,  5, -1 },
            {  0, -1,  0 }
        };
        int kernelSize = 3;
        int kernelOffset = kernelSize / 2;

        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                int r = 0, g = 0, b = 0;

                for (int ky = -kernelOffset; ky <= kernelOffset; ky++)
                {
                    int py = y + ky;
                    if (py < 0 || py >= height) continue;

                    for (int kx = -kernelOffset; kx <= kernelOffset; kx++)
                    {
                        int px = x + kx;
                        if (px < 0 || px >= width) continue;

                        int weight = kernel[ky + kernelOffset, kx + kernelOffset];
                        Color pixelColor = source.GetPixel(px, py);
                        r += pixelColor.R * weight;
                        g += pixelColor.G * weight;
                        b += pixelColor.B * weight;
                    }
                }

                // Clamp values to byte range.
                r = Math.Min(Math.Max(r, 0), 255);
                g = Math.Min(Math.Max(g, 0), 255);
                b = Math.Min(Math.Max(b, 0), 255);

                result.SetPixel(x, y, Color.FromArgb(r, g, b));
            }
        }

        return result;
    }
}

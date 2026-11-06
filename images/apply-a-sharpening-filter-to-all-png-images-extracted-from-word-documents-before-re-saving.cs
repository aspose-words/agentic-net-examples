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
        // Create a deterministic PNG image to be used in the document.
        const string sampleImagePath = "sample.png";
        CreateSamplePng(sampleImagePath);

        // Create a Word document and insert the PNG image.
        const string inputDocPath = "input.docx";
        CreateDocumentWithImage(sampleImagePath, inputDocPath);

        // Load the document for processing.
        Document doc = new Document(inputDocPath);

        // Process all PNG images in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
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
                    using (Bitmap sharpenedBitmap = ApplySharpeningFilter(originalBitmap))
                    {
                        // Save the processed bitmap to a new memory stream.
                        using (MemoryStream processedStream = new MemoryStream())
                        {
                            sharpenedBitmap.Save(processedStream, ImageFormat.Png);
                            processedStream.Position = 0;

                            // Replace the image in the shape with the processed one.
                            shape.ImageData.SetImage(processedStream);
                        }
                    }
                }
            }
        }

        // Save the modified document.
        const string outputDocPath = "output.docx";
        doc.Save(outputDocPath);

        // Validate that the output file was created.
        if (!File.Exists(outputDocPath))
            throw new InvalidOperationException("The output document was not created.");

        // Clean up sample files (optional).
        // File.Delete(sampleImagePath);
        // File.Delete(inputDocPath);
    }

    private static void CreateSamplePng(string filePath)
    {
        const int width = 100;
        const int height = 100;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a simple black rectangle.
                using (SolidBrush brush = new SolidBrush(Aspose.Drawing.Color.Black))
                {
                    g.FillRectangle(brush, 20, 20, 60, 60);
                }
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }

        if (!File.Exists(filePath))
            throw new InvalidOperationException("Failed to create the sample PNG image.");
    }

    private static void CreateDocumentWithImage(string imagePath, string docPath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        if (!File.Exists(docPath))
            throw new InvalidOperationException("Failed to create the input Word document.");
    }

    private static Bitmap ApplySharpeningFilter(Bitmap source)
    {
        int width = source.Width;
        int height = source.Height;
        Bitmap result = new Bitmap(width, height);

        // Sharpening kernel
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

                        int kernelValue = kernel[ky + offset, kx + offset];
                        Color pixelColor = source.GetPixel(px, py);
                        r += pixelColor.R * kernelValue;
                        g += pixelColor.G * kernelValue;
                        b += pixelColor.B * kernelValue;
                    }
                }

                // Clamp color values to byte range.
                r = Math.Min(Math.Max(r, 0), 255);
                g = Math.Min(Math.Max(g, 0), 255);
                b = Math.Min(Math.Max(b, 0), 255);

                result.SetPixel(x, y, Color.FromArgb(r, g, b));
            }
        }

        return result;
    }
}

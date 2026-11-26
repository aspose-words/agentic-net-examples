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
        // Create a deterministic sample PNG image.
        const string sampleImagePath = "sample.png";
        CreateSamplePng(sampleImagePath);

        // Create a Word document and insert the sample PNG image.
        const string inputDocPath = "input.docx";
        CreateDocumentWithImage(sampleImagePath, inputDocPath);

        // Load the document for processing.
        Document doc = new Document(inputDocPath);

        // Process all shapes that contain PNG images.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int pngCount = 0;
        int index = 0;

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Extract the original PNG image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the image into Aspose.Drawing.Bitmap.
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    // Apply sharpening filter.
                    using (Bitmap sharpenedBitmap = ApplySharpenFilter(originalBitmap))
                    {
                        // Save sharpened image to a new file for verification.
                        string sharpenedFileName = $"sharpened-{index}.png";
                        sharpenedBitmap.Save(sharpenedFileName, ImageFormat.Png);

                        // Replace the image in the document with the sharpened version.
                        using (MemoryStream sharpenedStream = new MemoryStream())
                        {
                            sharpenedBitmap.Save(sharpenedStream, ImageFormat.Png);
                            sharpenedStream.Position = 0;
                            shape.ImageData.SetImage(sharpenedStream);
                        }

                        pngCount++;
                    }
                }
            }

            index++;
        }

        // Validate that at least one PNG image was processed.
        if (pngCount == 0)
            throw new InvalidOperationException("No PNG images were found in the document.");

        // Save the modified document.
        const string outputDocPath = "output.docx";
        doc.Save(outputDocPath, SaveFormat.Docx);
    }

    private static void CreateSamplePng(string path)
    {
        const int width = 100;
        const int height = 100;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple black rectangle.
                using (SolidBrush brush = new SolidBrush(Color.Black))
                {
                    g.FillRectangle(brush, 20, 20, 60, 60);
                }
            }
            bitmap.Save(path, ImageFormat.Png);
        }
    }

    private static void CreateDocumentWithImage(string imagePath, string docPath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath, SaveFormat.Docx);
    }

    private static Bitmap ApplySharpenFilter(Bitmap source)
    {
        int width = source.Width;
        int height = source.Height;
        Bitmap result = new Bitmap(width, height);

        // Sharpen kernel.
        int[,] kernel = new int[,]
        {
            { 0, -1, 0 },
            { -1, 5, -1 },
            { 0, -1, 0 }
        };
        int kernelSize = 3;
        int offset = kernelSize / 2;

        // Process each pixel (excluding borders).
        for (int y = offset; y < height - offset; y++)
        {
            for (int x = offset; x < width - offset; x++)
            {
                int sumR = 0, sumG = 0, sumB = 0;

                for (int ky = -offset; ky <= offset; ky++)
                {
                    for (int kx = -offset; kx <= offset; kx++)
                    {
                        int pixelX = x + kx;
                        int pixelY = y + ky;
                        Color srcColor = source.GetPixel(pixelX, pixelY);
                        int weight = kernel[ky + offset, kx + offset];

                        sumR += srcColor.R * weight;
                        sumG += srcColor.G * weight;
                        sumB += srcColor.B * weight;
                    }
                }

                // Clamp values to byte range.
                sumR = Math.Max(0, Math.Min(255, sumR));
                sumG = Math.Max(0, Math.Min(255, sumG));
                sumB = Math.Max(0, Math.Min(255, sumB));

                result.SetPixel(x, y, Color.FromArgb(sumR, sumG, sumB));
            }
        }

        // Copy border pixels from the original image.
        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                if (x < offset || x >= width - offset || y < offset || y >= height - offset)
                {
                    result.SetPixel(x, y, source.GetPixel(x, y));
                }
            }
        }

        return result;
    }
}

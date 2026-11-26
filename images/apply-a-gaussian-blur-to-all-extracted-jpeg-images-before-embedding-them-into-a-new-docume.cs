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
        // Create a sample JPEG image.
        const string sampleImagePath = "sample.jpg";
        CreateSampleJpeg(sampleImagePath);

        // Create a source document and insert the sample JPEG.
        const string sourceDocPath = "source.docx";
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.InsertImage(sampleImagePath);
        sourceDoc.Save(sourceDocPath);

        // Load the source document for extraction.
        Document loadedSource = new Document(sourceDocPath);
        NodeCollection shapes = loadedSource.GetChildNodes(NodeType.Shape, true);

        // Prepare the output document.
        Document resultDoc = new Document();
        DocumentBuilder resultBuilder = new DocumentBuilder(resultDoc);

        int imageIndex = 0;
        int blurredCount = 0;

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Only process JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Extract the image to a memory stream.
            using (MemoryStream imgStream = new MemoryStream())
            {
                shape.ImageData.Save(imgStream);
                imgStream.Position = 0;

                // Load the image into a bitmap.
                using (Aspose.Drawing.Bitmap originalBitmap = new Aspose.Drawing.Bitmap(imgStream))
                {
                    // Apply Gaussian blur.
                    using (Aspose.Drawing.Bitmap blurredBitmap = ApplyGaussianBlur(originalBitmap))
                    {
                        // Save blurred image to a temporary file.
                        string blurredPath = $"blurred_{imageIndex}.jpg";
                        blurredBitmap.Save(blurredPath, ImageFormat.Jpeg);
                        blurredCount++;

                        // Insert blurred image into the result document.
                        resultBuilder.InsertImage(blurredPath);
                    }
                }
            }

            imageIndex++;
        }

        // Save the result document.
        const string resultDocPath = "result.docx";
        resultDoc.Save(resultDocPath);

        // Validation.
        if (!File.Exists(resultDocPath))
            throw new Exception("Result document was not created.");

        if (blurredCount == 0)
            throw new Exception("No JPEG images were processed and blurred.");

        // Clean up temporary files (optional).
        if (File.Exists(sampleImagePath)) File.Delete(sampleImagePath);
        if (File.Exists(sourceDocPath)) File.Delete(sourceDocPath);
        for (int i = 0; i < imageIndex; i++)
        {
            string blurredPath = $"blurred_{i}.jpg";
            if (File.Exists(blurredPath)) File.Delete(blurredPath);
        }
    }

    // Creates a simple deterministic JPEG image.
    private static void CreateSampleJpeg(string path)
    {
        const int width = 200;
        const int height = 100;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 3))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
                {
                    using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Red))
                    {
                        g.DrawString("Sample", font, brush, new Aspose.Drawing.PointF(30, 30));
                    }
                }
            }
            bitmap.Save(path, ImageFormat.Jpeg);
        }
    }

    // Applies a simple Gaussian blur using a 3x3 kernel.
    private static Aspose.Drawing.Bitmap ApplyGaussianBlur(Aspose.Drawing.Bitmap src)
    {
        int w = src.Width;
        int h = src.Height;
        Aspose.Drawing.Bitmap dst = new Aspose.Drawing.Bitmap(w, h);

        // 3x3 Gaussian kernel.
        int[,] kernel = {
            {1, 2, 1},
            {2, 4, 2},
            {1, 2, 1}
        };
        int kernelWeight = 16;

        for (int y = 0; y < h; y++)
        {
            for (int x = 0; x < w; x++)
            {
                int r = 0, g = 0, b = 0;

                for (int ky = -1; ky <= 1; ky++)
                {
                    int py = y + ky;
                    if (py < 0) py = 0;
                    if (py >= h) py = h - 1;

                    for (int kx = -1; kx <= 1; kx++)
                    {
                        int px = x + kx;
                        if (px < 0) px = 0;
                        if (px >= w) px = w - 1;

                        Aspose.Drawing.Color pixelColor = src.GetPixel(px, py);
                        int weight = kernel[ky + 1, kx + 1];

                        r += pixelColor.R * weight;
                        g += pixelColor.G * weight;
                        b += pixelColor.B * weight;
                    }
                }

                r = Math.Min(255, Math.Max(0, r / kernelWeight));
                g = Math.Min(255, Math.Max(0, g / kernelWeight));
                b = Math.Min(255, Math.Max(0, b / kernelWeight));

                dst.SetPixel(x, y, Aspose.Drawing.Color.FromArgb(r, g, b));
            }
        }

        return dst;
    }
}

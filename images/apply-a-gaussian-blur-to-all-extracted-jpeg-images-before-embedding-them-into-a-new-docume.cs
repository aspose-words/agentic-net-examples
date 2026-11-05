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

        // Load the source document and extract JPEG images.
        Document loadedSource = new Document(sourceDocPath);
        NodeCollection shapes = loadedSource.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage) continue;

            // Ensure the image is a JPEG.
            string imageFormat = shape.ImageData.ImageType.ToString();
            if (!imageFormat.Equals("Jpeg", StringComparison.OrdinalIgnoreCase))
                continue;

            // Extract image bytes.
            using (MemoryStream imgStream = new MemoryStream())
            {
                shape.ImageData.Save(imgStream);
                imgStream.Position = 0;

                // Load image into Aspose.Drawing.Bitmap.
                using (Bitmap originalBitmap = new Bitmap(imgStream))
                {
                    // Apply Gaussian blur.
                    using (Bitmap blurredBitmap = ApplyGaussianBlur(originalBitmap))
                    {
                        // Save blurred image to a temporary file.
                        string blurredPath = $"blurred_{imageIndex}.jpg";
                        blurredBitmap.Save(blurredPath, ImageFormat.Jpeg);

                        // Create a new document and embed the blurred image.
                        Document resultDoc = new Document();
                        DocumentBuilder resultBuilder = new DocumentBuilder(resultDoc);
                        resultBuilder.InsertImage(blurredPath);
                        string resultDocPath = $"result_{imageIndex}.docx";
                        resultDoc.Save(resultDocPath);

                        // Validate output.
                        if (!File.Exists(resultDocPath))
                            throw new InvalidOperationException($"Failed to create {resultDocPath}");
                        if (!File.Exists(blurredPath))
                            throw new InvalidOperationException($"Failed to create {blurredPath}");

                        // Clean up temporary blurred image file.
                        File.Delete(blurredPath);
                    }
                }
            }

            imageIndex++;
        }

        // Validate that at least one image was processed.
        if (imageIndex == 0)
            throw new InvalidOperationException("No JPEG images were found in the source document.");

        // Clean up sample files.
        File.Delete(sampleImagePath);
        File.Delete(sourceDocPath);
    }

    private static void CreateSampleJpeg(string path)
    {
        const int width = 200;
        const int height = 100;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                using (SolidBrush brush = new SolidBrush(Color.Red))
                {
                    g.FillRectangle(brush, 20, 20, 160, 60);
                }
            }
            bitmap.Save(path, ImageFormat.Jpeg);
        }
    }

    private static Bitmap ApplyGaussianBlur(Bitmap src)
    {
        int w = src.Width;
        int h = src.Height;
        Bitmap dst = new Bitmap(w, h);

        double[,] kernel = new double[5, 5]
        {
            { 1,  4,  7,  4, 1 },
            { 4, 16, 26, 16, 4 },
            { 7, 26, 41, 26, 7 },
            { 4, 16, 26, 16, 4 },
            { 1,  4,  7,  4, 1 }
        };
        double kernelSum = 273.0;

        for (int y = 0; y < h; y++)
        {
            for (int x = 0; x < w; x++)
            {
                double r = 0, g = 0, b = 0;
                for (int ky = -2; ky <= 2; ky++)
                {
                    int py = Math.Min(Math.Max(y + ky, 0), h - 1);
                    for (int kx = -2; kx <= 2; kx++)
                    {
                        int px = Math.Min(Math.Max(x + kx, 0), w - 1);
                        Color pixelColor = src.GetPixel(px, py);
                        double weight = kernel[ky + 2, kx + 2];
                        r += pixelColor.R * weight;
                        g += pixelColor.G * weight;
                        b += pixelColor.B * weight;
                    }
                }
                int ir = Math.Min(255, Math.Max(0, (int)(r / kernelSum)));
                int ig = Math.Min(255, Math.Max(0, (int)(g / kernelSum)));
                int ib = Math.Min(255, Math.Max(0, (int)(b / kernelSum)));
                dst.SetPixel(x, y, Color.FromArgb(ir, ig, ib));
            }
        }

        return dst;
    }
}

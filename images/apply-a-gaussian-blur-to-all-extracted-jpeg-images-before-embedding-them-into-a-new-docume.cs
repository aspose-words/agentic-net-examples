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
        // Step 1: Create a sample JPEG image.
        const string sampleImagePath = "sample.jpg";
        CreateSampleJpeg(sampleImagePath);

        // Step 2: Create a source document and insert the sample image.
        const string sourceDocPath = "source.docx";
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.InsertImage(sampleImagePath);
        sourceDoc.Save(sourceDocPath);

        // Step 3: Extract JPEG images from the source document.
        string[] extractedImages = ExtractJpegImages(sourceDoc, "extracted");

        if (extractedImages.Length == 0)
            throw new Exception("No JPEG images were extracted from the source document.");

        // Step 4: Apply Gaussian blur to each extracted image and collect blurred paths.
        string[] blurredImages = new string[extractedImages.Length];
        for (int i = 0; i < extractedImages.Length; i++)
        {
            string blurredPath = $"blurred_{i}.jpg";
            ApplyGaussianBlur(extractedImages[i], blurredPath);
            blurredImages[i] = blurredPath;
        }

        // Step 5: Create a new document and embed the blurred images.
        const string resultDocPath = "result.docx";
        Document resultDoc = new Document();
        DocumentBuilder resultBuilder = new DocumentBuilder(resultDoc);
        foreach (string blurredPath in blurredImages)
        {
            resultBuilder.InsertParagraph();
            resultBuilder.InsertImage(blurredPath);
        }
        resultDoc.Save(resultDocPath);

        // Validation
        if (!File.Exists(resultDocPath))
            throw new Exception("Result document was not created.");

        Console.WriteLine("Processing completed successfully.");
    }

    private static void CreateSampleJpeg(string path)
    {
        const int width = 200;
        const int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple red rectangle.
                using (SolidBrush brush = new SolidBrush(Color.Red))
                {
                    g.FillRectangle(brush, 50, 50, 100, 100);
                }
            }
            bitmap.Save(path, ImageFormat.Jpeg);
        }
    }

    private static string[] ExtractJpegImages(Document doc, string baseFileName)
    {
        var shapes = doc.GetChildNodes(NodeType.Shape, true);
        var extractedPaths = new System.Collections.Generic.List<string>();
        int index = 0;
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Jpeg)
            {
                string imagePath = $"{baseFileName}_{index}.jpg";
                shape.ImageData.Save(imagePath);
                extractedPaths.Add(imagePath);
                index++;
            }
        }
        return extractedPaths.ToArray();
    }

    private static void ApplyGaussianBlur(string inputPath, string outputPath)
    {
        using (Bitmap source = new Bitmap(inputPath))
        {
            int width = source.Width;
            int height = source.Height;
            using (Bitmap blurred = new Bitmap(width, height))
            {
                // Simple 5x5 Gaussian kernel (approximation).
                double[,] kernel = {
                    { 1,  4,  7,  4, 1 },
                    { 4, 16, 26, 16, 4 },
                    { 7, 26, 41, 26, 7 },
                    { 4, 16, 26, 16, 4 },
                    { 1,  4,  7,  4, 1 }
                };
                double kernelSum = 273; // Sum of all kernel values.

                for (int y = 0; y < height; y++)
                {
                    for (int x = 0; x < width; x++)
                    {
                        double r = 0, g = 0, b = 0;
                        for (int ky = -2; ky <= 2; ky++)
                        {
                            int py = Math.Min(height - 1, Math.Max(0, y + ky));
                            for (int kx = -2; kx <= 2; kx++)
                            {
                                int px = Math.Min(width - 1, Math.Max(0, x + kx));
                                Color pixelColor = source.GetPixel(px, py);
                                double weight = kernel[ky + 2, kx + 2];
                                r += pixelColor.R * weight;
                                g += pixelColor.G * weight;
                                b += pixelColor.B * weight;
                            }
                        }
                        int nr = Math.Min(255, Math.Max(0, (int)(r / kernelSum)));
                        int ng = Math.Min(255, Math.Max(0, (int)(g / kernelSum)));
                        int nb = Math.Min(255, Math.Max(0, (int)(b / kernelSum)));
                        blurred.SetPixel(x, y, Color.FromArgb(nr, ng, nb));
                    }
                }
                blurred.Save(outputPath, ImageFormat.Jpeg);
            }
        }
    }
}

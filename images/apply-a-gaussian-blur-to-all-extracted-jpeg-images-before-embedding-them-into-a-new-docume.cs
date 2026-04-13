using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string sampleImagePath = "sample.jpg";
        const string sourceDocPath = "source.docx";
        const string resultDocPath = "result.docx";

        // 1. Create a deterministic sample JPEG image.
        CreateSampleImage(sampleImagePath);

        // 2. Insert the sample image into a source document.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.InsertImage(sampleImagePath);
        sourceDoc.Save(sourceDocPath);
        ValidateFileExists(sourceDocPath, "source document");

        // 3. Load the source document and extract JPEG images.
        Document loadedDoc = new Document(sourceDocPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        var jpegShapes = shapeNodes
            .OfType<Shape>()
            .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Jpeg)
            .ToList();

        if (!jpegShapes.Any())
            throw new InvalidOperationException("No JPEG images were found in the source document.");

        // 4. Apply Gaussian blur to each extracted image and store blurred files.
        var blurredImagePaths = jpegShapes
            .Select((shape, index) => ApplyGaussianBlurToShape(shape, index))
            .ToList();

        // 5. Create a new document and embed the blurred images.
        Document resultDoc = new Document();
        DocumentBuilder resultBuilder = new DocumentBuilder(resultDoc);
        foreach (string blurredPath in blurredImagePaths)
        {
            resultBuilder.InsertImage(blurredPath);
        }
        resultDoc.Save(resultDocPath);
        ValidateFileExists(resultDocPath, "result document");

        // 6. Validate that blurred images were created.
        foreach (string path in blurredImagePaths)
            ValidateFileExists(path, "blurred image");
    }

    // Creates a simple 200x200 JPEG image with a red rectangle on a white background.
    private static void CreateSampleImage(string filePath)
    {
        const int width = 200;
        const int height = 200;
        Bitmap bitmap = new Bitmap(width, height);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        // Draw a solid red rectangle.
        using (SolidBrush brush = new SolidBrush(Color.Red))
        {
            graphics.FillRectangle(brush, 50, 50, 100, 100);
        }
        graphics.Dispose();
        bitmap.Save(filePath, ImageFormat.Jpeg);
        bitmap.Dispose();
    }

    // Applies Gaussian blur to the image contained in the given shape and saves it.
    private static string ApplyGaussianBlurToShape(Shape shape, int index)
    {
        // Retrieve image bytes from the shape.
        byte[] imageBytes = shape.ImageData.ImageBytes;
        using (MemoryStream ms = new MemoryStream(imageBytes))
        {
            ms.Position = 0;
            Bitmap original = new Bitmap(ms);
            Bitmap blurred = GaussianBlur(original);
            string blurredPath = $"blurred_{index}.jpg";
            blurred.Save(blurredPath, ImageFormat.Jpeg);
            original.Dispose();
            blurred.Dispose();
            return blurredPath;
        }
    }

    // Performs a simple Gaussian blur using a 5x5 kernel.
    private static Bitmap GaussianBlur(Bitmap source)
    {
        int width = source.Width;
        int height = source.Height;
        Bitmap result = new Bitmap(width, height);

        // 5x5 Gaussian kernel (sigma ≈ 1.0)
        double[,] kernel = {
            { 1,  4,  7,  4, 1 },
            { 4, 16, 26, 16, 4 },
            { 7, 26, 41, 26, 7 },
            { 4, 16, 26, 16, 4 },
            { 1,  4,  7,  4, 1 }
        };
        double kernelSum = 273.0; // Sum of all kernel values.

        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                double r = 0, g = 0, b = 0;
                for (int ky = -2; ky <= 2; ky++)
                {
                    int py = Math.Clamp(y + ky, 0, height - 1);
                    for (int kx = -2; kx <= 2; kx++)
                    {
                        int px = Math.Clamp(x + kx, 0, width - 1);
                        Color pixel = source.GetPixel(px, py);
                        double weight = kernel[ky + 2, kx + 2];
                        r += pixel.R * weight;
                        g += pixel.G * weight;
                        b += pixel.B * weight;
                    }
                }
                int ri = (int)Math.Round(r / kernelSum);
                int gi = (int)Math.Round(g / kernelSum);
                int bi = (int)Math.Round(b / kernelSum);
                result.SetPixel(x, y, Color.FromArgb(ri, gi, bi));
            }
        }

        return result;
    }

    // Simple helper to ensure a file was created.
    private static void ValidateFileExists(string path, string description)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException($"The expected {description} file was not created: {path}");
    }
}

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

        // 1. Create sample JPEG images.
        string[] sampleImagePaths = CreateSampleJpegImages(artifactsDir);

        // 2. Build a source document that contains the sample images.
        string sourceDocPath = Path.Combine(artifactsDir, "Source.docx");
        BuildSourceDocument(sampleImagePaths, sourceDocPath);

        // 3. Load the source document and apply Gaussian blur to each JPEG image.
        string outputDocPath = Path.Combine(artifactsDir, "Output.docx");
        ApplyGaussianBlurToJpegImages(sourceDocPath, outputDocPath);

        // 4. Validate that the output document was created.
        if (!File.Exists(outputDocPath))
            throw new InvalidOperationException("The output document was not created.");

        // (Optional) Clean up temporary blurred images.
        CleanupTemporaryFiles(artifactsDir);
    }

    // Creates a few deterministic JPEG images and returns their file paths.
    private static string[] CreateSampleJpegImages(string folder)
    {
        string[] paths = new string[2];

        for (int i = 0; i < paths.Length; i++)
        {
            int width = 200;
            int height = 200;
            using (Bitmap bitmap = new Bitmap(width, height))
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill background.
                g.Clear(Color.White);

                // Draw a colored rectangle.
                Color rectColor = i == 0 ? Color.Red : Color.Blue;
                using (Brush brush = new SolidBrush(rectColor))
                {
                    g.FillRectangle(brush, 20, 20, width - 40, height - 40);
                }

                // Save as JPEG.
                string filePath = Path.Combine(folder, $"Sample{i + 1}.jpg");
                bitmap.Save(filePath, ImageFormat.Jpeg);
                paths[i] = filePath;
            }
        }

        return paths;
    }

    // Inserts the provided images into a new document.
    private static void BuildSourceDocument(string[] imagePaths, string outputPath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        foreach (string imgPath in imagePaths)
        {
            builder.InsertParagraph();
            builder.InsertImage(imgPath);
        }

        doc.Save(outputPath);
    }

    // Loads a document, blurs each JPEG image, and saves the result.
    private static void ApplyGaussianBlurToJpegImages(string inputPath, string outputPath)
    {
        Document doc = new Document(inputPath);
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Extract image bytes.
            byte[] imageBytes = shape.ImageData.ImageBytes;
            using (MemoryStream ms = new MemoryStream(imageBytes))
            using (Bitmap bitmap = new Bitmap(ms))
            {
                // Apply a simple blur (average of 3x3 neighbourhood).
                ApplySimpleBlur(bitmap);

                // Save blurred image to a temporary stream.
                using (MemoryStream blurredStream = new MemoryStream())
                {
                    bitmap.Save(blurredStream, ImageFormat.Jpeg);
                    blurredStream.Position = 0;

                    // Replace the shape's image with the blurred version.
                    shape.ImageData.SetImage(blurredStream);
                }
            }

            imageIndex++;
        }

        doc.Save(outputPath);
    }

    // Simple 3x3 average blur (approximates Gaussian blur for demonstration).
    private static void ApplySimpleBlur(Bitmap source)
    {
        int width = source.Width;
        int height = source.Height;
        Bitmap temp = new Bitmap(width, height);

        for (int y = 1; y < height - 1; y++)
        {
            for (int x = 1; x < width - 1; x++)
            {
                int r = 0, g = 0, b = 0;
                for (int ky = -1; ky <= 1; ky++)
                {
                    for (int kx = -1; kx <= 1; kx++)
                    {
                        Color c = source.GetPixel(x + kx, y + ky);
                        r += c.R;
                        g += c.G;
                        b += c.B;
                    }
                }
                r /= 9;
                g /= 9;
                b /= 9;
                temp.SetPixel(x, y, Color.FromArgb(r, g, b));
            }
        }

        // Copy blurred pixels back to the original bitmap.
        for (int y = 1; y < height - 1; y++)
        {
            for (int x = 1; x < width - 1; x++)
            {
                source.SetPixel(x, y, temp.GetPixel(x, y));
            }
        }

        temp.Dispose();
    }

    // Removes any temporary blurred image files that might have been created.
    private static void CleanupTemporaryFiles(string folder)
    {
        foreach (string file in Directory.GetFiles(folder, "blurred_*.jpg"))
        {
            try { File.Delete(file); } catch { /* ignore */ }
        }
    }
}

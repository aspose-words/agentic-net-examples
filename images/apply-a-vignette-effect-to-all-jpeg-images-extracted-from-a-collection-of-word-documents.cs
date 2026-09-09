using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    // Entry point
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a sample JPEG image
        string sampleImagePath = Path.Combine(baseDir, "sample.jpg");
        CreateSampleJpeg(sampleImagePath, 200, 200);

        // Create sample Word documents containing the JPEG image
        for (int i = 1; i <= 2; i++)
        {
            string docPath = Path.Combine(inputDir, $"Document{i}.docx");
            CreateDocumentWithImage(docPath, sampleImagePath);
        }

        // Process each document: apply vignette to all JPEG images
        foreach (string docFile in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docFile);
            var shapeNodes = doc.GetChildNodes(NodeType.Shape, true)
                                .Cast<Shape>()
                                .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Jpeg)
                                .ToList();

            if (!shapeNodes.Any())
                throw new InvalidOperationException($"No JPEG images found in document '{docFile}'.");

            foreach (Shape shape in shapeNodes)
            {
                // Extract original JPEG image to a memory stream
                using (MemoryStream originalStream = new MemoryStream())
                {
                    shape.ImageData.Save(originalStream);
                    originalStream.Position = 0;

                    // Load bitmap from stream
                    using (Bitmap bitmap = new Bitmap(originalStream))
                    {
                        // Apply vignette effect
                        ApplyVignette(bitmap, 0.5f);

                        // Save processed bitmap back to a new stream as JPEG
                        using (MemoryStream processedStream = new MemoryStream())
                        {
                            bitmap.Save(processedStream, ImageFormat.Jpeg);
                            processedStream.Position = 0;

                            // Replace image in the shape
                            shape.ImageData.SetImage(processedStream);
                        }
                    }
                }
            }

            // Save the modified document
            string outputPath = Path.Combine(outputDir, Path.GetFileName(docFile));
            doc.Save(outputPath);
        }
    }

    // Creates a deterministic JPEG image using Aspose.Drawing
    private static void CreateSampleJpeg(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Aspose.Drawing.Color.LightBlue);
            // Draw a simple red ellipse in the center
            using (Brush brush = new SolidBrush(Aspose.Drawing.Color.Red))
            {
                int ellipseSize = Math.Min(width, height) / 2;
                int x = (width - ellipseSize) / 2;
                int y = (height - ellipseSize) / 2;
                g.FillEllipse(brush, x, y, ellipseSize, ellipseSize);
            }
            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }

    // Creates a Word document with the specified image inserted
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"Document containing image: {Path.GetFileName(imagePath)}");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Applies a simple vignette effect to the bitmap
    private static void ApplyVignette(Bitmap bitmap, float strength)
    {
        int width = bitmap.Width;
        int height = bitmap.Height;
        float centerX = width / 2f;
        float centerY = height / 2f;
        float maxDist = (float)Math.Sqrt(centerX * centerX + centerY * centerY);

        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                // Distance from center normalized [0,1]
                float dx = x - centerX;
                float dy = y - centerY;
                float dist = (float)Math.Sqrt(dx * dx + dy * dy);
                float factor = 1f - (dist / maxDist);
                factor = Math.Max(0f, factor);
                // Apply strength (the farther from center, the darker)
                float vignette = (float)Math.Pow(factor, 2) * (1 - strength) + strength;
                // Get original color
                Aspose.Drawing.Color orig = bitmap.GetPixel(x, y);
                // Apply vignette factor to each channel
                int r = (int)(orig.R * vignette);
                int g = (int)(orig.G * vignette);
                int b = (int)(orig.B * vignette);
                bitmap.SetPixel(x, y, Aspose.Drawing.Color.FromArgb(r, g, b));
            }
        }
    }
}

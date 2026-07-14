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
        // Directories for artifacts
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample JPEG image using Aspose.Drawing
        string originalImagePath = Path.Combine(artifactsDir, "sample.jpg");
        CreateSampleJpeg(originalImagePath, 200, 150);

        // 2. Create a Word document and insert the sample image
        string originalDocPath = Path.Combine(artifactsDir, "Original.docx");
        CreateDocumentWithImage(originalDocPath, originalImagePath);

        // 3. Load the document, extract JPEG images, apply blur, and re‑embed
        string blurredDocPath = Path.Combine(artifactsDir, "Blurred.docx");
        ApplyBlurToJpegImages(originalDocPath, blurredDocPath, artifactsDir);

        // Validation
        if (!File.Exists(blurredDocPath))
            throw new InvalidOperationException("The blurred document was not created.");

        Console.WriteLine("Processing completed successfully.");
    }

    // Creates a deterministic JPEG image.
    private static void CreateSampleJpeg(string filePath, int width, int height)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            g.Clear(Aspose.Drawing.Color.LightBlue);
            using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Red, 5))
            {
                g.DrawEllipse(pen, 20, 20, width - 40, height - 40);
            }
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
        }
    }

    // Creates a Word document containing the specified image.
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with original image:");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Applies a simple box blur to a bitmap (radius = 1).
    private static Aspose.Drawing.Bitmap ApplyBoxBlur(Aspose.Drawing.Bitmap source)
    {
        int width = source.Width;
        int height = source.Height;
        Aspose.Drawing.Bitmap blurred = new Aspose.Drawing.Bitmap(width, height);

        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                int rSum = 0, gSum = 0, bSum = 0, count = 0;

                // Iterate over the 3x3 neighbourhood
                for (int ny = y - 1; ny <= y + 1; ny++)
                {
                    for (int nx = x - 1; nx <= x + 1; nx++)
                    {
                        if (nx >= 0 && nx < width && ny >= 0 && ny < height)
                        {
                            Aspose.Drawing.Color c = source.GetPixel(nx, ny);
                            rSum += c.R;
                            gSum += c.G;
                            bSum += c.B;
                            count++;
                        }
                    }
                }

                // Average colour
                Aspose.Drawing.Color avg = Aspose.Drawing.Color.FromArgb(rSum / count, gSum / count, bSum / count);
                blurred.SetPixel(x, y, avg);
            }
        }

        return blurred;
    }

    // Loads a document, blurs each JPEG image, and saves a new document.
    private static void ApplyBlurToJpegImages(string sourceDocPath, string targetDocPath, string tempDir)
    {
        Document doc = new Document(sourceDocPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Extract image to a memory stream
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load image into Aspose.Drawing bitmap
                using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(originalStream))
                {
                    // Apply a simple box blur (radius 1)
                    using (Aspose.Drawing.Bitmap blurredBitmap = ApplyBoxBlur(bitmap))
                    {
                        // Save blurred image to a temporary file
                        string blurredImagePath = Path.Combine(tempDir, $"blurred_{imageIndex}.jpg");
                        blurredBitmap.Save(blurredImagePath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);

                        // Replace the shape's image with the blurred version
                        shape.ImageData.SetImage(blurredImagePath);
                    }
                }
            }

            imageIndex++;
        }

        // Save the modified document
        doc.Save(targetDocPath);
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Directory for all generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // ---------- 1. Create sample images ----------
        string imgPath1 = Path.Combine(outputDir, "sample1.png");
        string imgPath2 = Path.Combine(outputDir, "sample2.png");

        CreateSampleImage(imgPath1, 300, 200, Aspose.Drawing.Color.LightBlue, "Img1");
        CreateSampleImage(imgPath2, 200, 300, Aspose.Drawing.Color.LightCoral, "Img2");

        // ---------- 2. Create a simple HTML that references the images ----------
        string htmlPath = Path.Combine(outputDir, "sample.html");
        string htmlContent = $@"
<!DOCTYPE html>
<html>
<body>
    <h1>Test HTML</h1>
    <p>First image:</p>
    <img src=""{Path.GetFileName(imgPath1)}"" />
    <p>Second image:</p>
    <img src=""{Path.GetFileName(imgPath2)}"" />
</body>
</html>";
        File.WriteAllText(htmlPath, htmlContent);

        // ---------- 3. Load the HTML document ----------
        Document doc = new Document(htmlPath);

        // ---------- 4. Extract each image, generate a thumbnail while keeping aspect ratio ----------
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Save original image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0; // Reset before reading.

                // Load the image with Aspose.Drawing.
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    // Desired maximum thumbnail size.
                    const int maxThumbWidth = 100;
                    const int maxThumbHeight = 100;

                    // Compute scaling factor while preserving aspect ratio.
                    double widthScale = (double)maxThumbWidth / originalBitmap.Width;
                    double heightScale = (double)maxThumbHeight / originalBitmap.Height;
                    double scale = Math.Min(widthScale, heightScale);
                    if (scale > 1) scale = 1; // Do not upscale.

                    int thumbWidth = (int)(originalBitmap.Width * scale);
                    int thumbHeight = (int)(originalBitmap.Height * scale);

                    // Create thumbnail bitmap.
                    using (Bitmap thumbBitmap = new Bitmap(thumbWidth, thumbHeight))
                    {
                        using (Graphics g = Graphics.FromImage(thumbBitmap))
                        {
                            g.Clear(Aspose.Drawing.Color.White);
                            g.DrawImage(originalBitmap, 0, 0, thumbWidth, thumbHeight);
                        }

                        // Save thumbnail as PNG.
                        string thumbPath = Path.Combine(outputDir, $"thumb_{imageIndex}.png");
                        thumbBitmap.Save(thumbPath, Aspose.Drawing.Imaging.ImageFormat.Png);
                        if (!File.Exists(thumbPath))
                            throw new InvalidOperationException($"Thumbnail was not created: {thumbPath}");
                    }
                }
            }

            imageIndex++;
        }

        // Validation: at least one thumbnail must have been created.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the HTML document.");

        // Optional: clean up the temporary HTML file (not required).
        // File.Delete(htmlPath);
    }

    // Helper method to create a deterministic sample image.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backColor, string text)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(backColor);
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20))
                {
                    g.DrawString(text, font, Aspose.Drawing.Brushes.Black, new Aspose.Drawing.PointF(10, height / 2 - 10));
                }
            }
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create sample image: {filePath}");
    }
}

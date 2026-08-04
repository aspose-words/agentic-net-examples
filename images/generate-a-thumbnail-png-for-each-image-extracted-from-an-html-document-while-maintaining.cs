using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for all artifacts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // ---------- 1. Create sample images ----------
        string[] sampleImagePaths = { Path.Combine(artifactsDir, "sample1.png"), Path.Combine(artifactsDir, "sample2.jpg") };
        CreateSampleImage(sampleImagePaths[0], 300, 200, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(sampleImagePaths[1], 150, 250, Aspose.Drawing.Color.LightCoral);

        // ---------- 2. Build a simple HTML file that references the images ----------
        string htmlPath = Path.Combine(artifactsDir, "sample.html");
        string htmlContent = $@"
<!DOCTYPE html>
<html>
<head><title>Sample</title></head>
<body>
    <h1>Images</h1>
    <img src=""{sampleImagePaths[0]}"" alt=""Image 1""/>
    <p>Some text between images.</p>
    <img src=""{sampleImagePaths[1]}"" alt=""Image 2""/>
</body>
</html>";
        File.WriteAllText(htmlPath, htmlContent);

        // ---------- 3. Load the HTML document with Aspose.Words ----------
        Document doc = new Document(htmlPath);

        // ---------- 4. Extract each image, generate a thumbnail, and save it ----------
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Obtain the raw image bytes.
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load the image into Aspose.Drawing.Bitmap.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0;
                using (Bitmap originalBitmap = new Bitmap(ms))
                {
                    // Determine thumbnail size while preserving aspect ratio.
                    const int maxThumbSize = 100; // pixels
                    int thumbWidth, thumbHeight;
                    if (originalBitmap.Width >= originalBitmap.Height)
                    {
                        thumbWidth = maxThumbSize;
                        thumbHeight = (int)(originalBitmap.Height * (maxThumbSize / (double)originalBitmap.Width));
                    }
                    else
                    {
                        thumbHeight = maxThumbSize;
                        thumbWidth = (int)(originalBitmap.Width * (maxThumbSize / (double)originalBitmap.Height));
                    }

                    // Create a new bitmap for the thumbnail.
                    using (Bitmap thumbBitmap = new Bitmap(thumbWidth, thumbHeight))
                    {
                        using (Graphics g = Graphics.FromImage(thumbBitmap))
                        {
                            g.Clear(Aspose.Drawing.Color.White);
                            g.DrawImage(originalBitmap, 0, 0, thumbWidth, thumbHeight);
                        }

                        // Save the thumbnail as PNG.
                        string thumbPath = Path.Combine(artifactsDir, $"thumb_{imageIndex}.png");
                        thumbBitmap.Save(thumbPath, ImageFormat.Png);

                        // Validate that the file was created.
                        if (!File.Exists(thumbPath))
                            throw new InvalidOperationException($"Thumbnail was not saved: {thumbPath}");
                    }
                }
            }

            imageIndex++;
        }

        // If no images were found, raise an exception to satisfy validation rules.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the HTML document.");

        // All work is done; the program will exit automatically.
    }

    // Helper method to create a deterministic sample image.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(backColor);
                // Draw a simple rectangle border.
                using (Pen pen = new Pen(Aspose.Drawing.Color.Black, 3))
                {
                    g.DrawRectangle(pen, 0, 0, width - 1, height - 1);
                }
            }

            bitmap.Save(filePath);
        }
    }
}

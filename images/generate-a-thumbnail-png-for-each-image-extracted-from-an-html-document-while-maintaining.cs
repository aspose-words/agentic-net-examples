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
        // Prepare a folder for all artifacts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        if (Directory.Exists(artifactsDir))
            Directory.Delete(artifactsDir, true);
        Directory.CreateDirectory(artifactsDir);

        // ---------- Create sample images ----------
        string[] sampleImagePaths = new string[2];
        for (int i = 0; i < sampleImagePaths.Length; i++)
        {
            int width = 200 + i * 100;   // 200, 300
            int height = 150 + i * 50;   // 150, 200
            string imagePath = Path.Combine(artifactsDir, $"input{i + 1}.png");
            CreateSampleImage(width, height, imagePath);
            sampleImagePaths[i] = imagePath;
        }

        // ---------- Create a simple HTML file that references the images ----------
        string htmlPath = Path.Combine(artifactsDir, "sample.html");
        File.WriteAllText(htmlPath,
            $"<html><body>" +
            $"<p>First image:</p><img src=\"{Path.GetFileName(sampleImagePaths[0])}\"/>" +
            $"<p>Second image:</p><img src=\"{Path.GetFileName(sampleImagePaths[1])}\"/>" +
            $"</body></html>");

        // ---------- Load the HTML document ----------
        Document doc = new Document(htmlPath);

        // ---------- Extract each image, generate a thumbnail, and save it ----------
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Obtain the original image bytes.
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load the image into an Aspose.Drawing.Bitmap.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0;
                using (Bitmap original = new Bitmap(ms))
                {
                    // Determine thumbnail size while preserving aspect ratio (max 100x100).
                    const int maxSize = 100;
                    double ratio = Math.Min((double)maxSize / original.Width, (double)maxSize / original.Height);
                    if (ratio > 1) ratio = 1; // Do not upscale.

                    int thumbWidth = (int)(original.Width * ratio);
                    int thumbHeight = (int)(original.Height * ratio);

                    // Create the thumbnail bitmap.
                    using (Bitmap thumb = new Bitmap(thumbWidth, thumbHeight))
                    {
                        using (Graphics g = Graphics.FromImage(thumb))
                        {
                            g.Clear(Color.White);
                            g.DrawImage(original, 0, 0, thumbWidth, thumbHeight);
                        }

                        // Save the thumbnail as PNG.
                        string thumbPath = Path.Combine(artifactsDir, $"thumb_{imageIndex}.png");
                        thumb.Save(thumbPath, ImageFormat.Png);

                        // Validate that the thumbnail was created.
                        if (!File.Exists(thumbPath))
                            throw new InvalidOperationException($"Thumbnail not created: {thumbPath}");
                    }
                }
            }

            imageIndex++;
        }

        // Ensure at least one thumbnail was generated.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the HTML document.");

        // The program finishes automatically.
    }

    // Helper method to create a deterministic sample PNG image.
    private static void CreateSampleImage(int width, int height, string filePath)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle with a distinct color.
                using (Pen pen = new Pen(Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}

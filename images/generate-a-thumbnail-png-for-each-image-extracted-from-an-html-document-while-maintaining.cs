using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // ---------- Step 1: Create sample images ----------
        CreateSampleImage(Path.Combine(outputDir, "sample1.png"), 300, 200, Aspose.Drawing.Color.LightBlue, "Img 1");
        CreateSampleImage(Path.Combine(outputDir, "sample2.png"), 200, 300, Aspose.Drawing.Color.LightCoral, "Img 2");

        // ---------- Step 2: Create an HTML file referencing the images ----------
        string htmlContent = $@"
            <html>
                <body>
                    <p>First image:</p>
                    <img src=""{Path.GetFullPath(Path.Combine(outputDir, "sample1.png"))}"" />
                    <p>Second image:</p>
                    <img src=""{Path.GetFullPath(Path.Combine(outputDir, "sample2.png"))}"" />
                </body>
            </html>";
        string htmlPath = Path.Combine(outputDir, "sample.html");
        File.WriteAllText(htmlPath, htmlContent);

        // ---------- Step 3: Load HTML into Aspose.Words ----------
        LoadOptions loadOptions = new LoadOptions();
        Document doc = new Document(htmlPath, loadOptions);

        // ---------- Step 4: Extract images and generate thumbnails ----------
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage) continue;

            // Save the original extracted image.
            string extractedPath = Path.Combine(outputDir, $"extracted-{imageIndex}.png");
            shape.ImageData.Save(extractedPath);

            // Load the extracted image using Aspose.Drawing.
            using (Aspose.Drawing.Bitmap original = new Aspose.Drawing.Bitmap(extractedPath))
            {
                // Determine thumbnail size while preserving aspect ratio (max 100x100).
                const int maxThumbWidth = 100;
                const int maxThumbHeight = 100;
                double widthRatio = (double)maxThumbWidth / original.Width;
                double heightRatio = (double)maxThumbHeight / original.Height;
                double scale = Math.Min(widthRatio, heightRatio);
                int thumbWidth = (int)(original.Width * scale);
                int thumbHeight = (int)(original.Height * scale);

                // Create thumbnail bitmap.
                using (Aspose.Drawing.Bitmap thumbnail = new Aspose.Drawing.Bitmap(thumbWidth, thumbHeight))
                {
                    using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(thumbnail))
                    {
                        g.Clear(Aspose.Drawing.Color.White);
                        g.DrawImage(original, 0, 0, thumbWidth, thumbHeight);
                    }

                    // Save thumbnail as PNG.
                    string thumbPath = Path.Combine(outputDir, $"thumb-{imageIndex}.png");
                    thumbnail.Save(thumbPath, Aspose.Drawing.Imaging.ImageFormat.Png);

                    // Validate thumbnail creation.
                    if (!File.Exists(thumbPath))
                        throw new Exception($"Thumbnail was not created: {thumbPath}");
                }
            }

            imageIndex++;
        }

        // Validate that at least one thumbnail was generated.
        if (imageIndex == 0)
            throw new Exception("No images were found in the HTML document.");
    }

    // Helper method to create a deterministic sample image.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color bgColor, string text)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(bgColor);
                // Simple rectangle border.
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black, 3))
                {
                    g.DrawRectangle(pen, 0, 0, width - 1, height - 1);
                }
                // Draw text in the center.
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
                {
                    Aspose.Drawing.SizeF textSize = g.MeasureString(text, font);
                    float x = (width - textSize.Width) / 2;
                    float y = (height - textSize.Height) / 2;
                    g.DrawString(text, font, Aspose.Drawing.Brushes.Black, x, y);
                }
            }
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        // Validate that the sample image was created.
        if (!File.Exists(filePath))
            throw new Exception($"Sample image was not created: {filePath}");
    }
}

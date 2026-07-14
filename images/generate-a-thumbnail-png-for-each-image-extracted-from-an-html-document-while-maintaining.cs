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
        // Prepare a deterministic working folder.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // -----------------------------------------------------------------
        // 1. Create sample images that will be referenced from the HTML.
        // -----------------------------------------------------------------
        string imgPath1 = Path.Combine(workDir, "sample1.png");
        string imgPath2 = Path.Combine(workDir, "sample2.jpg");

        CreateSampleImage(imgPath1, 200, 120, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(imgPath2, 80, 150, Aspose.Drawing.Color.LightCoral);

        // -----------------------------------------------------------------
        // 2. Build a simple HTML document that contains the images.
        // -----------------------------------------------------------------
        string htmlPath = Path.Combine(workDir, "sample.html");
        string htmlContent = $@"
<html>
<body>
    <p>First image:</p>
    <img src=""{imgPath1}"" />
    <p>Second image:</p>
    <img src=""{imgPath2}"" />
</body>
</html>";
        File.WriteAllText(htmlPath, htmlContent);

        // -----------------------------------------------------------------
        // 3. Load the HTML into an Aspose.Words document.
        // -----------------------------------------------------------------
        Document doc = new Document(htmlPath);

        // -----------------------------------------------------------------
        // 4. Extract each image, generate a thumbnail while preserving aspect ratio,
        //    and save the thumbnail as PNG.
        // -----------------------------------------------------------------
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int thumbIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Obtain the raw image bytes from the shape.
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load the original image into an Aspose.Drawing.Bitmap.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0;
                using (Bitmap original = new Bitmap(ms))
                {
                    // Determine thumbnail size (max dimension = 100 pixels) while keeping aspect ratio.
                    const int maxDim = 100;
                    int thumbWidth, thumbHeight;

                    if (original.Width >= original.Height)
                    {
                        thumbWidth = maxDim;
                        thumbHeight = (int)Math.Round((double)original.Height / original.Width * maxDim);
                    }
                    else
                    {
                        thumbHeight = maxDim;
                        thumbWidth = (int)Math.Round((double)original.Width / original.Height * maxDim);
                    }

                    // Create the thumbnail bitmap.
                    using (Bitmap thumbnail = new Bitmap(thumbWidth, thumbHeight))
                    {
                        using (Graphics g = Graphics.FromImage(thumbnail))
                        {
                            // Clear background (optional) and draw the scaled image.
                            g.Clear(Aspose.Drawing.Color.White);
                            g.DrawImage(original, new Rectangle(0, 0, thumbWidth, thumbHeight));
                        }

                        // Save the thumbnail as PNG.
                        string thumbPath = Path.Combine(workDir, $"thumb_{thumbIndex}.png");
                        thumbnail.Save(thumbPath, ImageFormat.Png);
                        thumbIndex++;
                    }
                }
            }
        }

        // -----------------------------------------------------------------
        // 5. Validate that at least one thumbnail was created.
        // -----------------------------------------------------------------
        if (thumbIndex == 0)
            throw new InvalidOperationException("No images were extracted from the HTML document.");

        // The example finishes without requiring user interaction.
    }

    // Helper method to create a deterministic sample image.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(backColor);
                // Draw a simple diagonal line for visual distinction.
                using (Pen pen = new Pen(Aspose.Drawing.Color.Black, 2))
                {
                    g.DrawLine(pen, 0, 0, width - 1, height - 1);
                }
            }

            // Determine appropriate image format based on file extension.
            ImageFormat format = Path.GetExtension(filePath).Equals(".png", StringComparison.OrdinalIgnoreCase)
                ? ImageFormat.Png
                : ImageFormat.Jpeg;

            bitmap.Save(filePath, format);
        }
    }
}

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
        // Create deterministic sample images.
        CreateSampleImage("sample1.png", 200, 150, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage("sample2.png", 300, 100, Aspose.Drawing.Color.LightCoral);

        // Build simple HTML referencing the sample images.
        string htmlContent = @"
            <html>
                <body>
                    <p>First image:</p>
                    <img src='sample1.png' />
                    <p>Second image:</p>
                    <img src='sample2.png' />
                </body>
            </html>";

        // Save HTML to a local file.
        File.WriteAllText("sample.html", htmlContent);

        // Load the HTML document into Aspose.Words.
        Document doc = new Document("sample.html");

        // Extract images and generate thumbnails.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Save the original image to a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0;

                // Load the image into Aspose.Drawing.Bitmap.
                using (Bitmap originalBitmap = new Bitmap(imageStream))
                {
                    // Determine thumbnail size while preserving aspect ratio (max dimension 100).
                    const int maxDimension = 100;
                    double ratio = Math.Min((double)maxDimension / originalBitmap.Width, (double)maxDimension / originalBitmap.Height);
                    int thumbWidth = (int)(originalBitmap.Width * ratio);
                    int thumbHeight = (int)(originalBitmap.Height * ratio);
                    if (thumbWidth == 0) thumbWidth = 1;
                    if (thumbHeight == 0) thumbHeight = 1;

                    // Create thumbnail bitmap.
                    using (Bitmap thumbBitmap = new Bitmap(thumbWidth, thumbHeight))
                    {
                        using (Graphics graphics = Graphics.FromImage(thumbBitmap))
                        {
                            graphics.Clear(Aspose.Drawing.Color.White);
                            graphics.DrawImage(originalBitmap, 0, 0, thumbWidth, thumbHeight);
                        }

                        // Save thumbnail as PNG.
                        string thumbPath = $"thumb-{imageIndex}.png";
                        thumbBitmap.Save(thumbPath, ImageFormat.Png);

                        // Validate thumbnail creation.
                        if (!File.Exists(thumbPath))
                            throw new Exception($"Thumbnail was not created: {thumbPath}");
                    }
                }
            }

            imageIndex++;
        }

        // Ensure at least one thumbnail was generated.
        if (imageIndex == 0)
            throw new Exception("No images were extracted from the HTML document.");

        // Cleanup sample files (optional).
        // File.Delete("sample.html");
        // File.Delete("sample1.png");
        // File.Delete("sample2.png");
    }

    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color fillColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(fillColor);
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }

        // Validate image creation.
        if (!File.Exists(filePath))
            throw new Exception($"Failed to create sample image: {filePath}");
    }
}

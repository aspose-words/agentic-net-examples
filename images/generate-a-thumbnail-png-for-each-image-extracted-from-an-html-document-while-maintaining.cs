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
        // Step 1: Create sample images that will be referenced from HTML.
        CreateSampleImage("input1.png", Aspose.Drawing.Color.LightBlue);
        CreateSampleImage("input2.png", Aspose.Drawing.Color.LightCoral);

        // Step 2: Create a simple HTML file referencing the images.
        string htmlContent = @"
<html>
<body>
<p>First image:</p>
<img src='input1.png' />
<p>Second image:</p>
<img src='input2.png' />
</body>
</html>";
        File.WriteAllText("sample.html", htmlContent);

        // Step 3: Load the HTML document with Aspose.Words.
        Document doc = new Document("sample.html");

        // Step 4: Extract images from the document and generate thumbnails.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int thumbnailIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Save the original image to a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset before reading.

                // Load the image using Aspose.Drawing.
                using (Image originalImage = Image.FromStream(imageStream))
                {
                    // Determine thumbnail size while preserving aspect ratio (max 100x100).
                    const int maxSize = 100;
                    int originalWidth = originalImage.Width;
                    int originalHeight = originalImage.Height;
                    float ratio = Math.Min((float)maxSize / originalWidth, (float)maxSize / originalHeight);
                    int thumbWidth = (int)(originalWidth * ratio);
                    int thumbHeight = (int)(originalHeight * ratio);

                    // Create thumbnail bitmap.
                    using (Bitmap thumbBitmap = new Bitmap(thumbWidth, thumbHeight))
                    {
                        using (Graphics g = Graphics.FromImage(thumbBitmap))
                        {
                            g.Clear(Aspose.Drawing.Color.White);
                            g.DrawImage(originalImage, 0, 0, thumbWidth, thumbHeight);
                        }

                        // Save thumbnail as PNG.
                        string thumbFileName = $"thumb-{thumbnailIndex}.png";
                        thumbBitmap.Save(thumbFileName, ImageFormat.Png);
                        thumbnailIndex++;
                    }
                }
            }
        }

        // Validation: ensure at least one thumbnail was created.
        if (thumbnailIndex == 0)
            throw new InvalidOperationException("No images were extracted and thumbnails generated.");

        // Optional: clean up sample files (comment out if you want to inspect them).
        //File.Delete("input1.png");
        //File.Delete("input2.png");
        //File.Delete("sample.html");
    }

    private static void CreateSampleImage(string filePath, Aspose.Drawing.Color fillColor)
    {
        const int width = 200;
        const int height = 150;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(fillColor);
                // Draw a simple rectangle border.
                using (Pen pen = new Pen(Aspose.Drawing.Color.Black, 3))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}

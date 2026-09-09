using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image that will be referenced from HTML.
        // -----------------------------------------------------------------
        const int sampleWidth = 200;
        const int sampleHeight = 150;
        const string sampleImagePath = "input.png";

        // Use Aspose.Drawing types to avoid System.Drawing.
        using (Bitmap bmp = new Bitmap(sampleWidth, sampleHeight))
        using (Graphics g = Graphics.FromImage(bmp))
        {
            // Fill background with white and draw a simple rectangle.
            g.Clear(Color.White);
            using (SolidBrush brush = new SolidBrush(Color.LightBlue))
            {
                g.FillRectangle(brush, 10, 10, sampleWidth - 20, sampleHeight - 20);
            }

            // Save the sample image to disk.
            bmp.Save(sampleImagePath);
        }

        // -----------------------------------------------------------------
        // 2. Create a minimal HTML file that contains the image.
        // -----------------------------------------------------------------
        const string htmlPath = "sample.html";
        string htmlContent = $"<html><body><img src=\"{Path.GetFullPath(sampleImagePath)}\"/></body></html>";
        File.WriteAllText(htmlPath, htmlContent);

        // -----------------------------------------------------------------
        // 3. Load the HTML document with Aspose.Words.
        // -----------------------------------------------------------------
        Document doc = new Document(htmlPath);

        // -----------------------------------------------------------------
        // 4. Extract each image shape and generate a thumbnail PNG while
        //    preserving the original aspect ratio.
        // -----------------------------------------------------------------
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Save the shape's image to a memory stream.
            using (MemoryStream imgStream = new MemoryStream())
            {
                shape.ImageData.Save(imgStream);
                imgStream.Position = 0; // Reset before reading.

                // Load the original image using Aspose.Drawing.
                using (Bitmap original = new Bitmap(imgStream))
                {
                    // Determine thumbnail size (max 100x100) while keeping aspect ratio.
                    const int maxThumbSize = 100;
                    double scale = Math.Min((double)maxThumbSize / original.Width,
                                            (double)maxThumbSize / original.Height);
                    int thumbWidth = (int)Math.Round(original.Width * scale);
                    int thumbHeight = (int)Math.Round(original.Height * scale);

                    // Create the thumbnail bitmap.
                    using (Bitmap thumb = new Bitmap(thumbWidth, thumbHeight))
                    using (Graphics g = Graphics.FromImage(thumb))
                    {
                        g.Clear(Color.White);
                        // High‑quality scaling.
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.DrawImage(original, 0, 0, thumbWidth, thumbHeight);

                        // Save the thumbnail as PNG.
                        string thumbPath = $"thumb_{imageIndex}.png";
                        thumb.Save(thumbPath);
                        Console.WriteLine($"Thumbnail saved: {thumbPath}");
                    }
                }
            }

            imageIndex++;
        }

        // -----------------------------------------------------------------
        // 5. Validation – ensure at least one thumbnail was created.
        // -----------------------------------------------------------------
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the HTML document.");

        Console.WriteLine("Processing completed.");
    }
}

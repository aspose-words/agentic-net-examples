using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for all generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample JPEG image (200x200) using Aspose.Drawing.
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(artifactsDir, "sample.jpg");
        int originalWidth = 200;
        int originalHeight = 200;

        Bitmap bitmap = new Bitmap(originalWidth, originalHeight);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.LightBlue);
        // Draw a simple ellipse to make the image recognizable.
        graphics.DrawEllipse(Pens.DarkBlue, 20, 20, 160, 160);
        graphics.Dispose();
        bitmap.Save(sampleImagePath, ImageFormat.Jpeg);
        bitmap.Dispose();

        // -----------------------------------------------------------------
        // 2. Insert the image into a Word document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        string docPath = Path.Combine(artifactsDir, "DocumentWithImage.docx");
        doc.Save(docPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 3. Extract JPEG images, resize them to 50% and save as thumbnails.
        // -----------------------------------------------------------------
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int thumbnailIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Get the original image bytes.
            byte[] imageBytes = shape.ImageData.ToByteArray();

            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                // Load the original image.
                using (Bitmap originalBmp = new Bitmap(ms))
                {
                    // Calculate new dimensions (50% of original).
                    int thumbWidth = originalBmp.Width / 2;
                    int thumbHeight = originalBmp.Height / 2;

                    // Create a new bitmap for the thumbnail.
                    using (Bitmap thumbBmp = new Bitmap(thumbWidth, thumbHeight))
                    {
                        using (Graphics g = Graphics.FromImage(thumbBmp))
                        {
                            // High quality scaling.
                            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            g.DrawImage(originalBmp, 0, 0, thumbWidth, thumbHeight);
                        }

                        // Save the thumbnail.
                        string thumbPath = Path.Combine(artifactsDir, $"thumbnail_{thumbnailIndex}.jpg");
                        thumbBmp.Save(thumbPath, ImageFormat.Jpeg);
                        thumbnailIndex++;
                    }
                }
            }
        }

        // Validate that at least one thumbnail was created.
        if (thumbnailIndex == 0)
            throw new InvalidOperationException("No JPEG images were extracted and resized.");

        // Optional: indicate completion (no console interaction required).
    }
}

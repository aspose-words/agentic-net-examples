using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // File paths.
        string inputImagePath = Path.Combine(artifactsDir, "input.jpg");
        string docPath = Path.Combine(artifactsDir, "document.docx");
        string extractedImagePath = Path.Combine(artifactsDir, "extracted.jpg");
        string thumbnailPath = Path.Combine(artifactsDir, "thumbnail.jpg");

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample JPEG image using Aspose.Drawing.
        // -----------------------------------------------------------------
        int originalWidth = 800;
        int originalHeight = 600;
        using (Bitmap bitmap = new Bitmap(originalWidth, originalHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Simple visual content – a blue rectangle.
                using (SolidBrush brush = new SolidBrush(Color.Blue))
                {
                    g.FillRectangle(brush, 100, 100, 600, 400);
                }
            }
            // Save as JPEG.
            bitmap.Save(inputImagePath, ImageFormat.Jpeg);
        }

        // --------------------------------------------------------------
        // 2. Create a Word document and insert the JPEG image.
        // --------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape shape = builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // --------------------------------------------------------------
        // 3. Extract the image from the shape.
        // --------------------------------------------------------------
        if (!shape.HasImage)
            throw new InvalidOperationException("The inserted shape does not contain an image.");

        // Optional: save the extracted image directly.
        shape.ImageData.Save(extractedImagePath);

        // --------------------------------------------------------------
        // 4. Resize the extracted JPEG to 50% width and height.
        // --------------------------------------------------------------
        using (MemoryStream ms = new MemoryStream())
        {
            // Save image data to a stream and reset position.
            shape.ImageData.Save(ms);
            ms.Position = 0;

            using (Bitmap originalBitmap = new Bitmap(ms))
            {
                int thumbWidth = originalBitmap.Width / 2;
                int thumbHeight = originalBitmap.Height / 2;

                using (Bitmap thumbBitmap = new Bitmap(thumbWidth, thumbHeight))
                {
                    using (Graphics g = Graphics.FromImage(thumbBitmap))
                    {
                        // High‑quality scaling.
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.DrawImage(originalBitmap, 0, 0, thumbWidth, thumbHeight);
                    }
                    // Save the thumbnail as JPEG.
                    thumbBitmap.Save(thumbnailPath, ImageFormat.Jpeg);
                }
            }
        }

        // --------------------------------------------------------------
        // 5. Validate that the thumbnail file was created.
        // --------------------------------------------------------------
        if (!File.Exists(thumbnailPath))
            throw new FileNotFoundException("Thumbnail image was not created.", thumbnailPath);
    }
}

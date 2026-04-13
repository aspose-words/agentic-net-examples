using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D; // For InterpolationMode

public class Program
{
    public static void Main()
    {
        // Define deterministic file names.
        const string inputImagePath = "input.jpg";
        const string documentPath = "DocumentWithImage.docx";

        // -------------------------------------------------
        // 1. Create a sample JPEG image using Aspose.Drawing.
        // -------------------------------------------------
        const int originalWidth = 200;
        const int originalHeight = 200;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(originalWidth, originalHeight))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a simple red rectangle.
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Red, 5))
                {
                    g.DrawRectangle(pen, 10, 10, originalWidth - 20, originalHeight - 20);
                }
            }
            // Save as JPEG.
            bitmap.Save(inputImagePath, ImageFormat.Jpeg);
        }

        // Ensure the sample image exists.
        if (!File.Exists(inputImagePath))
            throw new FileNotFoundException("Failed to create the sample image.", inputImagePath);

        // -------------------------------------------------
        // 2. Create a Word document and insert the image.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape shape = builder.InsertImage(inputImagePath);
        // InsertImage already appends the shape to the document.
        doc.Save(documentPath);

        // -------------------------------------------------
        // 3. Extract JPEG images from the document, resize to 50%, and save as thumbnails.
        // -------------------------------------------------
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape imgShape in shapeNodes.OfType<Shape>())
        {
            if (!imgShape.HasImage)
                continue;

            // Process only JPEG images.
            if (imgShape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Save original image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                imgShape.ImageData.Save(originalStream);
                originalStream.Position = 0; // Reset before reading.

                // Load the image into Aspose.Drawing.Bitmap.
                using (Aspose.Drawing.Bitmap originalBitmap = new Aspose.Drawing.Bitmap(originalStream))
                {
                    // Calculate thumbnail size (50% of original).
                    int thumbWidth = originalBitmap.Width / 2;
                    int thumbHeight = originalBitmap.Height / 2;

                    // Create a new bitmap for the thumbnail.
                    using (Aspose.Drawing.Bitmap thumbBitmap = new Aspose.Drawing.Bitmap(thumbWidth, thumbHeight))
                    {
                        using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(thumbBitmap))
                        {
                            // High quality scaling.
                            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            g.DrawImage(originalBitmap, 0, 0, thumbWidth, thumbHeight);
                        }

                        // Save the thumbnail as JPEG.
                        string thumbPath = $"thumbnail_{imageIndex}.jpg";
                        thumbBitmap.Save(thumbPath, ImageFormat.Jpeg);

                        // Validate that the thumbnail was created.
                        if (!File.Exists(thumbPath))
                            throw new InvalidOperationException($"Thumbnail was not saved: {thumbPath}");
                    }
                }
            }

            imageIndex++;
        }

        // Ensure at least one thumbnail was produced.
        if (imageIndex == 0)
            throw new InvalidOperationException("No JPEG images were found to create thumbnails.");

        // Cleanup: optional removal of intermediate files.
        // File.Delete(inputImagePath);
        // File.Delete(documentPath);
    }
}

using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string sampleImagePath = "sample.jpg";
        const string documentPath = "document.docx";
        const string thumbnailPrefix = "thumb_";

        // 1. Create a deterministic sample JPEG image (200x200, solid blue)
        int originalWidth = 200;
        int originalHeight = 200;
        using (Bitmap bitmap = new Bitmap(originalWidth, originalHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.Blue);
            }
            // Save explicitly as JPEG
            bitmap.Save(sampleImagePath, ImageFormat.Jpeg);
        }

        // 2. Create a Word document and insert the sample image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(documentPath);

        // 3. Load the document (reuse the same instance is also fine)
        Document loadedDoc = new Document(documentPath);

        // 4. Extract JPEG images, resize them to 50% and save as thumbnails
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Obtain the image bytes
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load the image into Aspose.Drawing.Bitmap
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0;
                using (Bitmap originalBitmap = new Bitmap(ms))
                {
                    // Calculate thumbnail size (50% of original)
                    int thumbWidth = originalBitmap.Width / 2;
                    int thumbHeight = originalBitmap.Height / 2;

                    // Create thumbnail bitmap
                    using (Bitmap thumbBitmap = new Bitmap(thumbWidth, thumbHeight))
                    {
                        using (Graphics g = Graphics.FromImage(thumbBitmap))
                        {
                            // Draw the scaled image
                            g.DrawImage(
                                originalBitmap,
                                new Rectangle(0, 0, thumbWidth, thumbHeight));
                        }

                        // Save thumbnail explicitly as JPEG
                        string thumbPath = $"{thumbnailPrefix}{imageIndex}.jpg";
                        thumbBitmap.Save(thumbPath, ImageFormat.Jpeg);
                        imageIndex++;
                    }
                }
            }
        }

        // 5. Validation: ensure at least one thumbnail was created
        if (imageIndex == 0)
            throw new InvalidOperationException("No JPEG images were extracted and resized.");

        // Optional cleanup (commented out to keep outputs)
        // File.Delete(sampleImagePath);
        // File.Delete(documentPath);
    }
}

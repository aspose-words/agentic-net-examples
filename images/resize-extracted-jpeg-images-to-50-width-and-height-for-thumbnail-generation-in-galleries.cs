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
        // Step 1: Create a sample JPEG image.
        const string sampleImagePath = "sample.jpg";
        const int originalWidth = 800;
        const int originalHeight = 600;
        using (Bitmap bitmap = new Bitmap(originalWidth, originalHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.LightBlue);
                // Draw a simple rectangle.
                using (Pen pen = new Pen(Color.DarkBlue, 5))
                {
                    g.DrawRectangle(pen, 100, 100, 600, 400);
                }
            }
            bitmap.Save(sampleImagePath, ImageFormat.Jpeg);
        }

        // Step 2: Create a Word document and insert the JPEG image.
        const string docPath = "input.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // Step 3: Load the document and extract JPEG images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage) continue;
            if (shape.ImageData.ImageType != ImageType.Jpeg) continue;

            // Save the original image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the image using Aspose.Drawing.
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    int thumbWidth = originalBitmap.Width / 2;
                    int thumbHeight = originalBitmap.Height / 2;

                    // Create a thumbnail bitmap.
                    using (Bitmap thumbBitmap = new Bitmap(thumbWidth, thumbHeight))
                    {
                        using (Graphics g = Graphics.FromImage(thumbBitmap))
                        {
                            g.Clear(Color.White);
                            g.DrawImage(
                                originalBitmap,
                                new Rectangle(0, 0, thumbWidth, thumbHeight));
                        }

                        // Save the thumbnail.
                        string thumbPath = $"thumb-{imageIndex}.jpg";
                        thumbBitmap.Save(thumbPath, ImageFormat.Jpeg);

                        // Validate that the thumbnail was created.
                        if (!File.Exists(thumbPath))
                            throw new InvalidOperationException($"Thumbnail not created: {thumbPath}");
                    }
                }
            }

            imageIndex++;
        }

        // Final validation: ensure at least one thumbnail was generated.
        if (imageIndex == 0)
            throw new InvalidOperationException("No JPEG images were found to resize.");

        // Clean up sample files (optional).
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }
}

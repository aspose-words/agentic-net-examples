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
        const int originalWidth = 200;
        const int originalHeight = 200;

        using (Bitmap bitmap = new Bitmap(originalWidth, originalHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple red rectangle.
                using (Pen pen = new Pen(Color.Red, 5))
                {
                    g.DrawRectangle(pen, 10, 10, originalWidth - 20, originalHeight - 20);
                }
            }

            // Save as JPEG.
            bitmap.Save(sampleImagePath, ImageFormat.Jpeg);
        }

        // Step 2: Create a Word document and insert the image twice.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        builder.Writeln(); // Add a line break.
        builder.InsertImage(sampleImagePath);
        const string docPath = "DocumentWithImages.docx";
        doc.Save(docPath);

        // Step 3: Extract JPEG images from the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;
        int index = 0;
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                string extractedPath = $"extracted_{index}.jpg";
                shape.ImageData.Save(extractedPath);
                if (!File.Exists(extractedPath))
                    throw new InvalidOperationException($"Failed to save extracted image: {extractedPath}");

                // Step 4: Resize the extracted image to 50% for thumbnail.
                using (Bitmap originalBitmap = new Bitmap(extractedPath))
                {
                    int thumbWidth = originalBitmap.Width / 2;
                    int thumbHeight = originalBitmap.Height / 2;
                    using (Bitmap thumbBitmap = new Bitmap(thumbWidth, thumbHeight))
                    {
                        using (Graphics g = Graphics.FromImage(thumbBitmap))
                        {
                            g.Clear(Color.White);
                            // Draw the original image scaled down.
                            g.DrawImage(originalBitmap, 0, 0, thumbWidth, thumbHeight);
                        }

                        string thumbPath = $"thumb_{index}.jpg";
                        thumbBitmap.Save(thumbPath, ImageFormat.Jpeg);
                        if (!File.Exists(thumbPath))
                            throw new InvalidOperationException($"Failed to save thumbnail image: {thumbPath}");
                    }
                }

                extractedCount++;
                index++;
            }
        }

        // Validation: ensure at least one thumbnail was created.
        if (extractedCount == 0)
            throw new InvalidOperationException("No JPEG images were extracted from the document.");

        // Optional: clean up sample files (comment out if you want to keep them).
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }
}

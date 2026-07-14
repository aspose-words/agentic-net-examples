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
        const string sampleBmpPath = "sample.bmp";
        const string docPath = "docWithImage.docx";

        // 1. Create a sample BMP image (800x600) and save it locally
        const int originalWidth = 800;
        const int originalHeight = 600;
        using (Bitmap bitmap = new Bitmap(originalWidth, originalHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle for visual reference
                using (Pen pen = new Pen(Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 50, 50, originalWidth - 100, originalHeight - 100);
                }
            }
            bitmap.Save(sampleBmpPath, ImageFormat.Bmp);
        }

        // 2. Create a new document and insert the BMP image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleBmpPath);
        doc.Save(docPath);

        // 3. Load the document (already in memory) and extract images
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Save the original extracted image
            string originalImagePath = $"original_{imageIndex}.bmp";
            shape.ImageData.Save(originalImagePath);

            // Load the extracted image into a bitmap
            using (FileStream originalStream = File.OpenRead(originalImagePath))
            {
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    // Calculate new dimensions to have a width of 1024 pixels, preserving aspect ratio
                    const int targetWidth = 1024;
                    int targetHeight = (int)Math.Round((double)originalBitmap.Height * targetWidth / originalBitmap.Width);

                    // Create a new bitmap with the target size
                    using (Bitmap resizedBitmap = new Bitmap(targetWidth, targetHeight))
                    {
                        using (Graphics g = Graphics.FromImage(resizedBitmap))
                        {
                            g.Clear(Color.White);
                            // Draw the original image scaled to the new size
                            g.DrawImage(originalBitmap, 0, 0, targetWidth, targetHeight);
                        }

                        // Save the resized image
                        string resizedImagePath = $"resized_{imageIndex}.bmp";
                        resizedBitmap.Save(resizedImagePath, ImageFormat.Bmp);
                    }
                }
            }

            imageIndex++;
        }

        // 4. Validation: ensure at least one resized image was created
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted and resized.");

        // Cleanup temporary sample image (optional)
        if (File.Exists(sampleBmpPath))
            File.Delete(sampleBmpPath);
    }
}

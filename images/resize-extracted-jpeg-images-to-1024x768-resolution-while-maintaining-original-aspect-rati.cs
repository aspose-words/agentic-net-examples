using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample JPEG image.
        const string originalImagePath = "sample.jpg";
        const int originalWidth = 2000;
        const int originalHeight = 1500;
        using (Bitmap bitmap = new Bitmap(originalWidth, originalHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.LightBlue);
                // Draw a simple rectangle.
                using (Pen pen = new Pen(Color.DarkBlue, 10))
                {
                    g.DrawRectangle(pen, 100, 100, originalWidth - 200, originalHeight - 200);
                }
            }
            bitmap.Save(originalImagePath, ImageFormat.Jpeg);
        }

        // Step 2: Insert the image into a Word document.
        const string docPath = "DocumentWithImage.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(originalImagePath);
        doc.Save(docPath);

        // Step 3: Extract JPEG images, resize while maintaining aspect ratio, and save.
        const int maxWidth = 1024;
        const int maxHeight = 768;
        int imageIndex = 0;
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            if (!shape.HasImage)
                continue;

            // Ensure the image is JPEG.
            ImageType imageType = shape.ImageData.ImageType;
            if (imageType != ImageType.Jpeg)
                continue;

            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0;

                using (Bitmap originalBitmap = new Bitmap(imageStream))
                {
                    // Calculate new dimensions while preserving aspect ratio.
                    double ratioX = (double)maxWidth / originalBitmap.Width;
                    double ratioY = (double)maxHeight / originalBitmap.Height;
                    double ratio = Math.Min(ratioX, ratioY);
                    int newWidth = (int)(originalBitmap.Width * ratio);
                    int newHeight = (int)(originalBitmap.Height * ratio);

                    using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                    {
                        using (Graphics graphics = Graphics.FromImage(resizedBitmap))
                        {
                            graphics.Clear(Color.White);
                            graphics.DrawImage(originalBitmap, 0, 0, newWidth, newHeight);
                        }

                        string resizedImagePath = $"resized-{imageIndex}.jpg";
                        resizedBitmap.Save(resizedImagePath, ImageFormat.Jpeg);

                        if (!File.Exists(resizedImagePath))
                            throw new InvalidOperationException($"Failed to create resized image: {resizedImagePath}");
                    }
                }
            }

            imageIndex++;
        }

        // Validation: ensure at least one image was processed.
        if (imageIndex == 0)
            throw new InvalidOperationException("No JPEG images were found to resize.");

        // Cleanup: optional removal of intermediate files (commented out).
        // File.Delete(originalImagePath);
        // File.Delete(docPath);
    }
}

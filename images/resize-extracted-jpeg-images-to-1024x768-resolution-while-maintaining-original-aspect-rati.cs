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
        // Create a deterministic sample JPEG image.
        const string inputImagePath = "sample.jpg";
        using (Bitmap bitmap = new Bitmap(2000, 1500))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle for visual content.
                g.FillRectangle(new SolidBrush(Color.LightBlue), 100, 100, 1800, 1300);
            }
            bitmap.Save(inputImagePath, ImageFormat.Jpeg);
        }

        // Create a Word document and insert the sample image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // Load the document (optional, demonstrates load/save lifecycle).
        Document loadedDoc = new Document(docPath);

        // Extract JPEG images, resize them while maintaining aspect ratio, and save.
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int resizedCount = 0;

        for (int i = 0; i < shapes.Count; i++)
        {
            Shape shape = (Shape)shapes[i];
            if (!shape.HasImage)
                continue;

            // Process only JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Save the original image to a memory stream.
            using (MemoryStream ms = new MemoryStream())
            {
                shape.ImageData.Save(ms);
                ms.Position = 0; // Reset stream before reading.

                // Load the image into a bitmap.
                using (Bitmap original = new Bitmap(ms))
                {
                    // Determine new size within 1024x768 while preserving aspect ratio.
                    const int maxWidth = 1024;
                    const int maxHeight = 768;
                    double widthScale = (double)maxWidth / original.Width;
                    double heightScale = (double)maxHeight / original.Height;
                    double scale = Math.Min(widthScale, heightScale);
                    if (scale > 1) scale = 1; // Do not upscale if image is smaller.

                    int newWidth = (int)(original.Width * scale);
                    int newHeight = (int)(original.Height * scale);

                    // Resize the image.
                    using (Bitmap resized = new Bitmap(newWidth, newHeight))
                    {
                        using (Graphics g = Graphics.FromImage(resized))
                        {
                            g.DrawImage(original, 0, 0, newWidth, newHeight);
                        }

                        // Save the resized image.
                        string outputPath = $"resized-{i + 1}.jpg";
                        resized.Save(outputPath, ImageFormat.Jpeg);
                        resizedCount++;
                    }
                }
            }
        }

        // Validate that at least one image was resized and saved.
        if (resizedCount == 0)
            throw new InvalidOperationException("No JPEG images were found and resized.");

        // Optional cleanup: delete temporary files if desired.
        // File.Delete(inputImagePath);
        // File.Delete(docPath);
    }
}

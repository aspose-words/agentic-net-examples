using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample PNG image (200x200) and save it as input.png
        const int originalWidth = 200;
        const int originalHeight = 200;
        string inputImagePath = "input.png";

        using (Bitmap bitmap = new Bitmap(originalWidth, originalHeight))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Aspose.Drawing.Color.White);
            using (Brush brush = new SolidBrush(Aspose.Drawing.Color.Red))
            {
                g.FillRectangle(brush, 50, 50, 100, 100);
            }
            bitmap.Save(inputImagePath);
        }

        // Create a Word document and insert the sample image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape shape = builder.InsertImage(inputImagePath);
        // Ensure the shape is appended to the document (InsertImage already does this)

        // Prepare to extract and resize PNG images
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int previewIndex = 0;
        foreach (Shape imgShape in shapes.OfType<Shape>())
        {
            if (!imgShape.HasImage)
                continue;

            // Process only PNG images
            if (imgShape.ImageData.ImageType != ImageType.Png)
                continue;

            // Obtain the original image bytes
            byte[] imageBytes = imgShape.ImageData.ToByteArray();

            using (MemoryStream originalStream = new MemoryStream(imageBytes))
            {
                originalStream.Position = 0; // Reset stream position before use
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    // Calculate 75% of original dimensions
                    int newWidth = (int)(originalBitmap.Width * 0.75);
                    int newHeight = (int)(originalBitmap.Height * 0.75);

                    // Create a new bitmap for the resized preview
                    using (Bitmap previewBitmap = new Bitmap(newWidth, newHeight))
                    using (Graphics graphics = Graphics.FromImage(previewBitmap))
                    {
                        graphics.Clear(Aspose.Drawing.Color.Transparent);
                        // Draw the original image scaled to the new size
                        graphics.DrawImage(originalBitmap, 0, 0, newWidth, newHeight);

                        // Save the preview image
                        string previewPath = $"preview_{previewIndex}.png";
                        previewBitmap.Save(previewPath, ImageFormat.Png);
                        previewIndex++;
                    }
                }
            }
        }

        // Validate that at least one preview image was created
        if (previewIndex == 0)
            throw new InvalidOperationException("No PNG images were extracted and resized.");

        // Optional: Save the document (not required by the task but demonstrates full workflow)
        doc.Save("DocumentWithImage.docx");
    }
}

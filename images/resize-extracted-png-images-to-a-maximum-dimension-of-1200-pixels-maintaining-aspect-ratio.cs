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
        // Create a deterministic sample PNG image.
        const string inputImagePath = "input.png";
        const int sampleWidth = 2000;
        const int sampleHeight = 1500;
        using (Bitmap bmp = new Bitmap(sampleWidth, sampleHeight))
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.Clear(Aspose.Drawing.Color.LightBlue);
            // Draw a simple rectangle for visual reference.
            using (Pen pen = new Pen(Aspose.Drawing.Color.DarkBlue, 10))
            {
                g.DrawRectangle(pen, 100, 100, sampleWidth - 200, sampleHeight - 200);
            }
            bmp.Save(inputImagePath, ImageFormat.Png);
        }

        // Insert the sample image into a Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        const string docPath = "DocumentWithImage.docx";
        doc.Save(docPath);

        // Load the document (demonstrating load rule usage).
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Obtain the image bytes from the shape.
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load the bytes into an Aspose.Drawing.Bitmap.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0;
                using (Bitmap originalBitmap = new Bitmap(ms))
                {
                    // Determine new dimensions while preserving aspect ratio.
                    int originalWidth = originalBitmap.Width;
                    int originalHeight = originalBitmap.Height;
                    const int maxDimension = 1200;

                    double scale = 1.0;
                    if (originalWidth > originalHeight && originalWidth > maxDimension)
                        scale = (double)maxDimension / originalWidth;
                    else if (originalHeight >= originalWidth && originalHeight > maxDimension)
                        scale = (double)maxDimension / originalHeight;

                    int newWidth = (int)Math.Round(originalWidth * scale);
                    int newHeight = (int)Math.Round(originalHeight * scale);

                    // If scaling is not required, keep original size.
                    if (scale >= 1.0)
                    {
                        newWidth = originalWidth;
                        newHeight = originalHeight;
                    }

                    // Resize the image.
                    using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                    using (Graphics graphics = Graphics.FromImage(resizedBitmap))
                    {
                        graphics.Clear(Aspose.Drawing.Color.Transparent);
                        graphics.DrawImage(originalBitmap, 0, 0, newWidth, newHeight);

                        // Save the resized image.
                        string resizedImagePath = $"resized_{imageIndex}.png";
                        resizedBitmap.Save(resizedImagePath, ImageFormat.Png);

                        // Validate that the file was created.
                        if (!File.Exists(resizedImagePath))
                            throw new InvalidOperationException($"Failed to create resized image: {resizedImagePath}");
                    }
                }
            }

            imageIndex++;
        }

        // Ensure at least one image was processed.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Cleanup temporary files (optional).
        // File.Delete(inputImagePath);
        // File.Delete(docPath);
    }
}

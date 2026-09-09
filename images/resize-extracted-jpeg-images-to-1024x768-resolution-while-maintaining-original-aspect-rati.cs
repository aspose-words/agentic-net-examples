using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a sample JPEG image (1500x1000) to work with.
        const string inputImagePath = "sample_input.jpg";
        CreateSampleJpeg(inputImagePath, 1500, 1000);

        // Create a new document and insert the sample image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        const string docPath = "sample_doc.docx";
        doc.Save(docPath);

        // Extract JPEG images from the document, resize them to fit within 1024x768
        // while preserving aspect ratio, and save the resized versions.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;
            if (shape.ImageData.ImageType != ImageType.Jpeg) continue;

            // Save the original image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the image using Aspose.Drawing.
                using (Image originalImage = Image.FromStream(originalStream))
                {
                    // Determine new dimensions while maintaining aspect ratio.
                    const int maxWidth = 1024;
                    const int maxHeight = 768;
                    int newWidth = originalImage.Width;
                    int newHeight = originalImage.Height;

                    double widthRatio = (double)maxWidth / originalImage.Width;
                    double heightRatio = (double)maxHeight / originalImage.Height;
                    double scale = Math.Min(1.0, Math.Min(widthRatio, heightRatio));

                    newWidth = (int)(originalImage.Width * scale);
                    newHeight = (int)(originalImage.Height * scale);

                    // Resize the image.
                    using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                    {
                        using (Graphics g = Graphics.FromImage(resizedBitmap))
                        {
                            g.Clear(Color.White);
                            g.DrawImage(originalImage, 0, 0, newWidth, newHeight);
                        }

                        // Save the resized image.
                        string resizedPath = $"resized_image_{imageIndex}.jpg";
                        resizedBitmap.Save(resizedPath, ImageFormat.Jpeg);
                        if (!File.Exists(resizedPath))
                            throw new InvalidOperationException($"Failed to create resized image: {resizedPath}");
                    }
                }
            }

            imageIndex++;
        }

        if (imageIndex == 0)
            throw new InvalidOperationException("No JPEG images were found to resize.");
    }

    private static void CreateSampleJpeg(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.LightBlue);
                // Draw a simple rectangle for visual reference.
                using (Pen pen = new Pen(Color.DarkBlue, 5))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }

            bitmap.Save(filePath, ImageFormat.Jpeg);
        }

        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create sample image: {filePath}");
    }
}

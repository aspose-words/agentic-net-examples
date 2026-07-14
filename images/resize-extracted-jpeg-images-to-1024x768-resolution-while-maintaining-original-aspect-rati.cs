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
        // Create a sample JPEG image (2000x1500) to work with.
        const string inputImagePath = "input.jpg";
        CreateSampleJpeg(inputImagePath, 2000, 1500);

        // Create a new document and insert the sample image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);

        // Extract JPEG images from the document, resize them to fit within 1024x768
        // while preserving aspect ratio, and save the resized versions.
        int resizedCount = ExtractAndResizeJpegs(doc, 1024, 768);

        // Validate that at least one image was processed.
        if (resizedCount == 0)
            throw new Exception("No JPEG images were extracted from the document.");
    }

    private static void CreateSampleJpeg(string filePath, int width, int height)
    {
        // Create a bitmap, fill it with white, and save as JPEG.
        Bitmap bitmap = new Bitmap(width, height);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        // Dispose drawing objects before saving.
        graphics.Dispose();
        bitmap.Save(filePath, ImageFormat.Jpeg);
        bitmap.Dispose();
    }

    private static int ExtractAndResizeJpegs(Document doc, int maxWidth, int maxHeight)
    {
        int index = 0;
        // Get all shape nodes in the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage || shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Save the image data to a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset before reading.

                // Load the original image.
                using (Image original = Image.FromStream(imageStream))
                {
                    int originalWidth = original.Width;
                    int originalHeight = original.Height;

                    // Determine scaling factor to fit within the target dimensions.
                    double widthRatio = (double)maxWidth / originalWidth;
                    double heightRatio = (double)maxHeight / originalHeight;
                    double scale = Math.Min(widthRatio, heightRatio);
                    // If the image is already smaller, keep original size.
                    if (scale > 1.0) scale = 1.0;

                    int newWidth = (int)(originalWidth * scale);
                    int newHeight = (int)(originalHeight * scale);

                    // Resize the image.
                    using (Bitmap resized = new Bitmap(newWidth, newHeight))
                    {
                        using (Graphics g = Graphics.FromImage(resized))
                        {
                            g.DrawImage(original, 0, 0, newWidth, newHeight);
                        }

                        string outputPath = $"resized_{index}.jpg";
                        resized.Save(outputPath, ImageFormat.Jpeg);
                    }
                }
            }

            index++;
        }

        return index;
    }
}

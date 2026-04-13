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
        // Create a deterministic sample PNG image (800x400) and save it as input.png
        const string inputImagePath = "input.png";
        using (Bitmap bitmap = new Bitmap(800, 400))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle to make the image non‑empty
                g.FillRectangle(new SolidBrush(Color.Blue), 100, 100, 600, 200);
            }
            bitmap.Save(inputImagePath, ImageFormat.Png);
        }

        // Create a Word document and insert the sample image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        const string docPath = "doc_with_image.docx";
        doc.Save(docPath);

        // Extract PNG images from the document, resize them to a height of 600px while preserving aspect ratio,
        // and save the resized images as resized_0.png, resized_1.png, etc.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int resizedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage || shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Save the original image bytes to a memory stream
            using (MemoryStream ms = new MemoryStream())
            {
                shape.ImageData.Save(ms);
                ms.Position = 0; // Reset position before reading

                // Load the image into Aspose.Drawing.Bitmap
                using (Bitmap original = new Bitmap(ms))
                {
                    const int targetHeight = 600;
                    double scale = (double)targetHeight / original.Height;
                    int targetWidth = (int)(original.Width * scale);

                    // Create a new bitmap with the target dimensions
                    using (Bitmap resized = new Bitmap(targetWidth, targetHeight))
                    {
                        using (Graphics g = Graphics.FromImage(resized))
                        {
                            g.Clear(Color.White);
                            g.DrawImage(original, 0, 0, targetWidth, targetHeight);
                        }

                        string outputPath = $"resized_{resizedCount}.png";
                        resized.Save(outputPath, ImageFormat.Png);

                        // Validate that the file was created
                        if (!File.Exists(outputPath))
                            throw new InvalidOperationException($"Failed to create resized image: {outputPath}");

                        resizedCount++;
                    }
                }
            }
        }

        // Ensure at least one image was processed
        if (resizedCount == 0)
            throw new InvalidOperationException("No PNG images were found and resized in the document.");
    }
}

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
        // Create a deterministic PNG image to be used as input.
        const string inputImagePath = "input.png";
        CreateSamplePng(inputImagePath, 200, 200);

        // Build a document and insert the PNG image multiple times.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        builder.InsertParagraph();
        builder.InsertImage(inputImagePath);
        const string docPath = "DocWithImages.docx";
        doc.Save(docPath);

        // Extract all PNG images, apply a 5‑pixel red border, and save them.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;
            if (shape.ImageData.ImageType != ImageType.Png) continue;

            // Obtain the raw image bytes.
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load the bytes into a non‑indexed bitmap, draw the original image onto it,
            // then add a red border.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0;
                using (Bitmap original = new Bitmap(ms))
                {
                    // Create a new bitmap with a 32‑bpp ARGB pixel format (non‑indexed).
                    using (Bitmap bitmap = new Bitmap(original.Width, original.Height, PixelFormat.Format32bppArgb))
                    {
                        // Copy the original image onto the new bitmap.
                        using (Graphics g = Graphics.FromImage(bitmap))
                        {
                            g.DrawImage(original, 0, 0, original.Width, original.Height);

                            // Draw a 5‑pixel red border.
                            using (Pen pen = new Pen(Color.Red, 5))
                            {
                                // Adjust rectangle to stay inside the bitmap bounds.
                                g.DrawRectangle(pen, 0, 0, bitmap.Width - 1, bitmap.Height - 1);
                            }
                        }

                        // Save the modified image.
                        string outputPath = $"extracted_{extractedCount}.png";
                        bitmap.Save(outputPath, ImageFormat.Png);
                        extractedCount++;
                    }
                }
            }
        }

        // Validation: ensure at least one image was processed.
        if (extractedCount == 0)
            throw new InvalidOperationException("No PNG images were extracted from the document.");
    }

    // Helper method to create a simple PNG image using Aspose.Drawing.
    private static void CreateSamplePng(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple black ellipse for visual content.
                using (Pen pen = new Pen(Color.Black, 2))
                {
                    g.DrawEllipse(pen, 10, 10, width - 20, height - 20);
                }
            }

            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}

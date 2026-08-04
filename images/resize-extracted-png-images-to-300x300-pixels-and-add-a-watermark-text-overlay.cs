using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing; // Provides Bitmap, Graphics, Font, Color, etc.

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a deterministic sample PNG image (500x500) with text.
        // -----------------------------------------------------------------
        const string inputImagePath = "input.png";

        // Create bitmap and graphics objects from Aspose.Drawing.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(500, 500);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);

        // Draw sample text.
        using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 48))
        {
            graphics.DrawString(
                "Sample",
                font,
                new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black),
                new Aspose.Drawing.PointF(100, 200));
        }

        // Save the bitmap to a local file.
        bitmap.Save(inputImagePath);

        // Clean up drawing resources.
        graphics.Dispose();
        bitmap.Dispose();

        // -----------------------------------------------------------------
        // 2. Create a Word document and insert the sample image.
        // -----------------------------------------------------------------
        const string docPath = "doc.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Extract PNG images, resize to 300x300, add watermark, and save.
        // -----------------------------------------------------------------
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Png)
            {
                // Save the shape's image to a memory stream.
                using (MemoryStream imageStream = new MemoryStream())
                {
                    shape.ImageData.Save(imageStream);
                    imageStream.Position = 0; // Reset before reading.

                    // Load the extracted image into a bitmap.
                    using (Aspose.Drawing.Bitmap originalBitmap = new Aspose.Drawing.Bitmap(imageStream))
                    {
                        // Create a new 300x300 bitmap.
                        using (Aspose.Drawing.Bitmap resizedBitmap = new Aspose.Drawing.Bitmap(300, 300))
                        {
                            // Draw the original image scaled to 300x300.
                            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(resizedBitmap))
                            {
                                g.DrawImage(
                                    originalBitmap,
                                    new Aspose.Drawing.RectangleF(0, 0, 300, 300));

                                // Add watermark text overlay.
                                using (Aspose.Drawing.Font watermarkFont = new Aspose.Drawing.Font("Arial", 24))
                                using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(
                                    Aspose.Drawing.Color.FromArgb(128, Aspose.Drawing.Color.Red)))
                                {
                                    string watermarkText = "WATERMARK";
                                    // Position near the bottom‑left corner.
                                    g.DrawString(watermarkText, watermarkFont, brush, new Aspose.Drawing.PointF(10, 260));
                                }
                            }

                            // Save the watermarked image.
                            string outputPath = $"output_{imageIndex}.png";
                            resizedBitmap.Save(outputPath);

                            // Validate that the file was created.
                            if (!File.Exists(outputPath))
                                throw new Exception($"Failed to create output image: {outputPath}");
                        }
                    }
                }

                imageIndex++;
            }
        }

        // Ensure at least one image was processed.
        if (imageIndex == 0)
            throw new Exception("No PNG images were found in the document.");
    }
}

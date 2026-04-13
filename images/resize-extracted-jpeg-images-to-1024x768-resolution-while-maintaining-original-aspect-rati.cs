using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample JPEG image (1600x1200) using Aspose.Drawing.
        const string sampleImagePath = "sample.jpg";
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(1600, 1200))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.LightBlue);
                // Draw a simple rectangle to make the image non‑blank.
                g.FillRectangle(Aspose.Drawing.Brushes.Coral, 200, 150, 1200, 900);
            }
            bitmap.Save(sampleImagePath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
        }

        // Create a Word document and insert the sample image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape insertedShape = builder.InsertImage(sampleImagePath);
        // Ensure the shape is appended to the document (InsertImage already does this).
        doc.Save("document.docx");

        // Reload the document to simulate extraction scenario.
        Document loadedDoc = new Document("document.docx");
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        int resizedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Save the original image extracted from the shape.
            string extractedPath = $"extracted_{extractedCount}.jpg";
            shape.ImageData.Save(extractedPath);
            extractedCount++;

            // Load the extracted image using Aspose.Drawing.
            using (Aspose.Drawing.Image originalImage = Aspose.Drawing.Image.FromFile(extractedPath))
            {
                int originalWidth = originalImage.Width;
                int originalHeight = originalImage.Height;

                // Compute scaling factor to fit within 1024x768 while preserving aspect ratio.
                double scale = Math.Min(1024.0 / originalWidth, 768.0 / originalHeight);
                // If the image is already smaller than the target size, keep original dimensions.
                if (scale > 1.0) scale = 1.0;

                int newWidth = (int)(originalWidth * scale);
                int newHeight = (int)(originalHeight * scale);

                // Resize the image.
                using (Aspose.Drawing.Bitmap resizedBitmap = new Aspose.Drawing.Bitmap(newWidth, newHeight))
                {
                    using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(resizedBitmap))
                    {
                        graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        graphics.DrawImage(originalImage, 0, 0, newWidth, newHeight);
                    }

                    string resizedPath = $"resized_{resizedCount}.jpg";
                    resizedBitmap.Save(resizedPath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
                    resizedCount++;
                }
            }
        }

        // Validation: ensure at least one resized image was produced.
        if (resizedCount == 0)
            throw new InvalidOperationException("No JPEG images were resized.");

        // Optional cleanup (commented out).
        // File.Delete(sampleImagePath);
        // File.Delete("document.docx");
    }
}

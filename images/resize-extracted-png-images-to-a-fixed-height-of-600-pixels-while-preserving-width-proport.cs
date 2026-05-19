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
        // Prepare directories
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample PNG image (800x400) using Aspose.Drawing
        string inputImagePath = Path.Combine(artifactsDir, "input.png");
        using (Bitmap bitmap = new Bitmap(800, 400))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle for visual reference
                g.FillRectangle(new SolidBrush(Color.Blue), 0, 0, 800, 400);
            }
            bitmap.Save(inputImagePath, ImageFormat.Png);
        }

        // 2. Create a new Word document and insert the sample image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        string docPath = Path.Combine(artifactsDir, "DocumentWithImage.docx");
        doc.Save(docPath);

        // 3. Extract PNG images from the document
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            string extractedPath = Path.Combine(artifactsDir, $"extracted_{extractedCount}.png");
            shape.ImageData.Save(extractedPath);
            extractedCount++;

            // 4. Resize the extracted PNG to a fixed height of 600 pixels while preserving aspect ratio
            using (Bitmap originalBitmap = new Bitmap(extractedPath))
            {
                int originalWidth = originalBitmap.Width;
                int originalHeight = originalBitmap.Height;

                // Desired height
                int targetHeight = 600;
                // Compute proportional width
                int targetWidth = (int)Math.Round((double)originalWidth * targetHeight / originalHeight);

                using (Bitmap resizedBitmap = new Bitmap(targetWidth, targetHeight))
                {
                    using (Graphics graphics = Graphics.FromImage(resizedBitmap))
                    {
                        graphics.Clear(Color.Transparent);
                        graphics.DrawImage(originalBitmap, 0, 0, targetWidth, targetHeight);
                    }
                    string resizedPath = Path.Combine(artifactsDir, $"resized_{extractedCount - 1}.png");
                    resizedBitmap.Save(resizedPath, ImageFormat.Png);
                }
            }
        }

        // 5. Validation: ensure at least one resized image was created
        if (extractedCount == 0)
            throw new InvalidOperationException("No PNG images were extracted from the document.");

        int resizedImages = Directory.GetFiles(artifactsDir, "resized_*.png").Length;
        if (resizedImages == 0)
            throw new InvalidOperationException("Resizing failed: no resized PNG images were produced.");

        // Example completed successfully
        Console.WriteLine($"Processed {extractedCount} PNG image(s). Resized images are saved in: {artifactsDir}");
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;   // For InterpolationMode

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        string inputImagePath = "input.png";
        string docPath = "document.docx";

        // 1. Create a sample PNG image (200x200) and save it.
        int originalWidth = 200;
        int originalHeight = 200;
        using (Bitmap bitmap = new Bitmap(originalWidth, originalHeight))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            // Draw a simple red rectangle for visual distinction
            using (Pen pen = new Pen(Color.Red, 5))
            {
                graphics.DrawRectangle(pen, 10, 10, originalWidth - 20, originalHeight - 20);
            }
            bitmap.Save(inputImagePath);
        }

        // 2. Create a new Word document and insert the PNG image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // 3. Load the document (simulating a separate load step).
        Document loadedDoc = new Document(docPath);

        // 4. Extract each PNG image, resize to 75% and save as a preview.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Save the original image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0; // Reset before reading.

                // Load the original image using Aspose.Drawing.
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    // Calculate new dimensions (75% of original).
                    int newWidth = (int)(originalBitmap.Width * 0.75);
                    int newHeight = (int)(originalBitmap.Height * 0.75);

                    // Create a new bitmap for the resized image.
                    using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                    using (Graphics graphics = Graphics.FromImage(resizedBitmap))
                    {
                        // High‑quality scaling.
                        graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        graphics.DrawImage(
                            originalBitmap,
                            new Rectangle(0, 0, newWidth, newHeight));

                        // Save the resized preview.
                        string previewPath = $"preview_{imageIndex}.png";
                        resizedBitmap.Save(previewPath);
                        if (!File.Exists(previewPath))
                            throw new InvalidOperationException($"Failed to create preview image: {previewPath}");
                    }
                }
            }

            imageIndex++;
        }

        // Validate that at least one preview was generated.
        if (imageIndex == 0)
            throw new InvalidOperationException("No PNG images were found to generate previews.");

        // Optional cleanup (commented out to keep files for inspection).
        // File.Delete(inputImagePath);
        // File.Delete(docPath);
    }
}

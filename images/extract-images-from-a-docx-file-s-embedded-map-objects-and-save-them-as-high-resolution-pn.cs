using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a deterministic high‑resolution PNG that will act as the map image.
        const string mapImagePath = "map.png";
        const int width = 800;
        const int height = 600;

        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                graphics.Clear(Aspose.Drawing.Color.White);
                // Draw a simple rectangle to visualize the image.
                graphics.DrawRectangle(
                    new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5),
                    50, 50, width - 100, height - 100);
            }

            // Save the image as PNG.
            bitmap.Save(mapImagePath);
        }

        // Create a new Word document and insert the map image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(mapImagePath);
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // Reload the document to simulate extraction from an existing file.
        Document loadedDoc = new Document(docPath);

        // Collect all Shape nodes that contain images.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                // Generate a deterministic file name for each extracted image.
                string extractedImagePath = $"extracted-image-{imageIndex}.png";

                // Save the image data as PNG. The original image is already PNG,
                // so the format is preserved.
                shape.ImageData.Save(extractedImagePath);

                // Validate that the file was created.
                if (!File.Exists(extractedImagePath))
                {
                    throw new InvalidOperationException($"Failed to save extracted image to '{extractedImagePath}'.");
                }

                extractedCount++;
                imageIndex++;
            }
        }

        // Ensure at least one image was extracted.
        if (extractedCount == 0)
        {
            throw new InvalidOperationException("No images were extracted from the document.");
        }

        // Clean up temporary files (optional).
        // File.Delete(mapImagePath);
        // File.Delete(docPath);
    }
}

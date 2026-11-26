using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample image that will act as a map.
        const string mapImagePath = "map.png";
        CreateSampleMapImage(mapImagePath);

        // Create a DOCX document and insert the sample map image.
        const string docPath = "sample.docx";
        CreateDocumentWithMapImage(docPath, mapImagePath);

        // Load the document and extract all images from shape nodes.
        ExtractImagesFromDocument(docPath);
    }

    private static void CreateSampleMapImage(string filePath)
    {
        // Create a high‑resolution bitmap.
        using (Bitmap bitmap = new Bitmap(1200, 800))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                graphics.Clear(Color.White);

                // Draw a simple map representation.
                using (Pen pen = new Pen(Color.Blue, 8))
                {
                    graphics.DrawRectangle(pen, 100, 100, 1000, 600);
                }

                using (Pen pen = new Pen(Color.Red, 6))
                {
                    graphics.DrawLine(pen, 100, 100, 1100, 700);
                    graphics.DrawLine(pen, 1100, 100, 100, 700);
                }
            }

            // Save the bitmap as PNG.
            bitmap.Save(filePath, ImageFormat.Png);
        }

        // Validate that the image file was created.
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create sample image at '{filePath}'.");
    }

    private static void CreateDocumentWithMapImage(string docPath, string imagePath)
    {
        // Ensure the image exists before insertion.
        if (!File.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");

        // Create a new document and insert the image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);

        // Save the document.
        doc.Save(docPath);

        // Validate that the document file was created.
        if (!File.Exists(docPath))
            throw new InvalidOperationException($"Failed to create document at '{docPath}'.");
    }

    private static void ExtractImagesFromDocument(string docPath)
    {
        // Load the document.
        Document doc = new Document(docPath);

        // Get all shape nodes in the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            // Process only shapes that contain an image.
            if (shape.HasImage)
            {
                string outputFileName = $"extracted-{extractedCount + 1}.png";

                // Save the image data as PNG. Using the overload that accepts only the file name.
                shape.ImageData.Save(outputFileName);

                // Validate that the image file was written.
                if (!File.Exists(outputFileName))
                    throw new InvalidOperationException($"Failed to save extracted image to '{outputFileName}'.");

                extractedCount++;
            }
        }

        // Ensure at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");
    }
}

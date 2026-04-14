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
        // Directories for artifacts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a deterministic PNG sample image.
        string pngPath = Path.Combine(artifactsDir, "input.png");
        CreateSamplePng(pngPath, 200, 200);

        // 2. Build a Word document and insert the PNG image.
        string docPath = Path.Combine(artifactsDir, "document.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(pngPath);
        doc.Save(docPath);

        // 3. Extract each image, convert PNG to JPEG while preserving dimensions.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Save the original image (expected PNG) to a temporary file.
            string extractedPng = Path.Combine(artifactsDir, $"extracted_{imageIndex}.png");
            shape.ImageData.Save(extractedPng);

            // Load the PNG with Aspose.Drawing and save as JPEG.
            using (Image img = Image.FromFile(extractedPng))
            {
                string jpegPath = Path.Combine(artifactsDir, $"extracted_{imageIndex}.jpg");
                img.Save(jpegPath, ImageFormat.Jpeg);

                // Validate that the JPEG file was created.
                if (!File.Exists(jpegPath))
                    throw new InvalidOperationException($"Failed to create JPEG file: {jpegPath}");
            }

            imageIndex++;
        }

        // Ensure at least one image was processed.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");
    }

    // Helper method to create a simple PNG image using Aspose.Drawing.
    private static void CreateSamplePng(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                g.Clear(Aspose.Drawing.Color.White);

                // Draw a red rectangle for visual content.
                using (Pen pen = new Pen(Aspose.Drawing.Color.Red, 5))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }

            // Save the bitmap as PNG.
            bitmap.Save(filePath, ImageFormat.Png);
        }

        // Verify that the PNG file exists.
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create sample PNG image: {filePath}");
    }
}

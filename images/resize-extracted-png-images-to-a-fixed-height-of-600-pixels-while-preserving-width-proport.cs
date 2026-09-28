using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample PNG image.
        const string inputImagePath = "input.png";
        CreateSamplePng(inputImagePath, 400, 300);

        // Create a Word document and insert the sample image.
        const string docPath = "sample.docx";
        CreateDocumentWithImage(docPath, inputImagePath);

        // Load the document and process PNG images.
        Document doc = new Document(docPath);
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Extract the original PNG.
            string extractedPath = $"extracted-{imageIndex}.png";
            using (FileStream fs = new FileStream(extractedPath, FileMode.Create, FileAccess.Write))
            {
                shape.ImageData.Save(fs);
                fs.Flush();
            }

            // Validate extraction.
            if (!File.Exists(extractedPath))
                throw new InvalidOperationException($"Failed to extract image to {extractedPath}");

            // Resize the extracted PNG to a fixed height of 600 pixels while preserving aspect ratio.
            string resizedPath = $"resized-{imageIndex}.png";
            ResizePngToHeight(extractedPath, resizedPath, 600);

            // Validate resizing.
            if (!File.Exists(resizedPath))
                throw new InvalidOperationException($"Failed to save resized image to {resizedPath}");

            imageIndex++;
        }

        // Ensure at least one image was processed.
        if (imageIndex == 0)
            throw new InvalidOperationException("No PNG images were found in the document.");

        // Optional cleanup can be added here.
    }

    private static void CreateSamplePng(string path, int width, int height)
    {
        // Create a bitmap using Aspose.Drawing.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap);
        g.Clear(Aspose.Drawing.Color.White);

        // Draw a simple red rectangle for visual distinction.
        using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Red))
        {
            g.FillRectangle(brush, width / 4, height / 4, width / 2, height / 2);
        }

        // Save the bitmap as PNG.
        bitmap.Save(path, ImageFormat.Png);

        // Dispose resources.
        g.Dispose();
        bitmap.Dispose();

        if (!File.Exists(path))
            throw new InvalidOperationException($"Failed to create sample image at {path}");
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
        if (!File.Exists(docPath))
            throw new InvalidOperationException($"Failed to save document at {docPath}");
    }

    private static void ResizePngToHeight(string sourcePath, string destPath, int targetHeight)
    {
        // Load the original image.
        using (Aspose.Drawing.Bitmap original = new Aspose.Drawing.Bitmap(sourcePath))
        {
            int originalWidth = original.Width;
            int originalHeight = original.Height;
            if (originalHeight == 0)
                throw new InvalidOperationException("Original image height is zero.");

            double scaleFactor = (double)targetHeight / originalHeight;
            int targetWidth = (int)Math.Round(originalWidth * scaleFactor);

            // Create a new bitmap with the target dimensions.
            using (Aspose.Drawing.Bitmap resized = new Aspose.Drawing.Bitmap(targetWidth, targetHeight))
            {
                using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(resized))
                {
                    // High quality resizing settings.
                    g.InterpolationMode = Aspose.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.DrawImage(original, 0, 0, targetWidth, targetHeight);
                }

                // Save the resized image as PNG.
                resized.Save(destPath, ImageFormat.Png);
            }
        }
    }
}

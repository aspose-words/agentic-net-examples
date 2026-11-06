using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a deterministic sample PNG image.
        const string sampleImagePath = "sample.png";
        CreateSamplePng(sampleImagePath, 200, 200);

        // Step 2: Create a Word document and insert the sample image.
        const string docPath = "DocumentWithImage.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // Step 3: Load the document and extract PNG images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage) continue;

            // Save the original image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0; // Reset before reading.

                // Load the image with Aspose.Drawing.
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    // Calculate 50% dimensions.
                    int newWidth = originalBitmap.Width / 2;
                    int newHeight = originalBitmap.Height / 2;

                    // Create a new bitmap for the resized preview.
                    using (Bitmap previewBitmap = new Bitmap(newWidth, newHeight))
                    {
                        using (Graphics graphics = Graphics.FromImage(previewBitmap))
                        {
                            // Draw the original image scaled down.
                            graphics.DrawImage(originalBitmap, 0, 0, newWidth, newHeight);
                        }

                        // Save the preview image.
                        string previewPath = $"preview-{imageIndex}.png";
                        previewBitmap.Save(previewPath, ImageFormat.Png);

                        // Validate that the preview file was created.
                        if (!File.Exists(previewPath))
                            throw new InvalidOperationException($"Failed to create preview image: {previewPath}");

                        Console.WriteLine($"Preview image saved: {previewPath}");
                    }
                }
            }

            imageIndex++;
        }

        // Ensure at least one image was processed.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were found in the document.");

        Console.WriteLine("Image resizing completed successfully.");
    }

    private static void CreateSamplePng(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                graphics.Clear(Color.White);
                // Draw a simple red rectangle.
                using (Pen pen = new Pen(Color.Red, 5))
                {
                    graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }

            // Save as PNG.
            bitmap.Save(filePath, ImageFormat.Png);
        }

        // Validate that the sample image was created.
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create sample image: {filePath}");
    }
}

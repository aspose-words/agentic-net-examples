using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Text;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample PNG image that will be inserted into the Word documents.
        const string sampleImagePath = "sample.png";
        CreateSamplePng(sampleImagePath);

        // Step 2: Create a few Word documents and insert the sample PNG into each.
        string[] docPaths = { "Doc1.docx", "Doc2.docx" };
        for (int i = 0; i < docPaths.Length; i++)
        {
            CreateWordDocumentWithImage(docPaths[i], sampleImagePath);
        }

        // Step 3: Extract PNG images from each document, apply a watermark, and save the result.
        for (int docIndex = 0; docIndex < docPaths.Length; docIndex++)
        {
            string docPath = docPaths[docIndex];
            Document doc = new Document(docPath);

            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapes)
            {
                if (!shape.HasImage)
                    continue;

                // Only process PNG images.
                if (shape.ImageData.ImageType != ImageType.Png)
                    continue;

                // Extract the PNG image.
                string extractedImagePath = $"extracted-{docIndex + 1}-{imageIndex + 1}.png";
                shape.ImageData.Save(extractedImagePath);

                // Apply watermark to the extracted image.
                string watermarkedImagePath = $"watermarked-{docIndex + 1}-{imageIndex + 1}.png";
                ApplyWatermarkToPng(extractedImagePath, watermarkedImagePath, "WATERMARK");

                // Validate that the watermarked file was created.
                if (!File.Exists(watermarkedImagePath))
                    throw new InvalidOperationException($"Failed to create watermarked image: {watermarkedImagePath}");

                imageIndex++;
            }

            // Ensure at least one image was processed for the current document.
            if (imageIndex == 0)
                throw new InvalidOperationException($"No PNG images found in document: {docPath}");
        }

        Console.WriteLine("Watermarking completed successfully.");
    }

    private static void CreateSamplePng(string filePath)
    {
        const int width = 200;
        const int height = 100;

        // Create bitmap using Aspose.Drawing.
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            // Obtain graphics object.
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                g.Clear(Aspose.Drawing.Color.White);

                // Draw a simple red rectangle.
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Red, 3))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }

            // Save as PNG.
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        // Validate that the sample image was created.
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create sample image: {filePath}");
    }

    private static void CreateWordDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the PNG image into the document.
        builder.InsertImage(imagePath);

        // Save the document.
        doc.Save(docPath);
    }

    private static void ApplyWatermarkToPng(string inputPath, string outputPath, string watermarkText)
    {
        // Load the original PNG.
        using (Aspose.Drawing.Bitmap original = new Aspose.Drawing.Bitmap(inputPath))
        {
            // Create a new bitmap with the same dimensions.
            using (Aspose.Drawing.Bitmap watermarked = new Aspose.Drawing.Bitmap(original.Width, original.Height))
            {
                using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(watermarked))
                {
                    // Draw the original image onto the new bitmap.
                    g.DrawImage(original, 0, 0, original.Width, original.Height);

                    // Prepare watermark drawing tools.
                    using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20, Aspose.Drawing.FontStyle.Bold))
                    using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(
                        Aspose.Drawing.Color.FromArgb(128, Aspose.Drawing.Color.Yellow)))
                    {
                        // Measure the size of the watermark text.
                        SizeF textSize = g.MeasureString(watermarkText, font);

                        // Position the watermark at the bottom-right corner with a small margin.
                        float x = watermarked.Width - textSize.Width - 10;
                        float y = watermarked.Height - textSize.Height - 10;

                        // Draw the watermark text.
                        g.DrawString(watermarkText, font, brush, new PointF(x, y));
                    }
                }

                // Save the watermarked image as PNG.
                watermarked.Save(outputPath, Aspose.Drawing.Imaging.ImageFormat.Png);
            }
        }

        // Validate that the watermarked image was saved.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Watermarked image not saved: {outputPath}");
    }
}

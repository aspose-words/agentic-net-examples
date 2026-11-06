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
        const string inputImagePath = "input.png";
        CreateSamplePng(inputImagePath);

        // Step 2: Create a Word document and insert the sample image.
        const string docPath = "sample.docx";
        CreateDocumentWithImage(docPath, inputImagePath);

        // Step 3: Load the document, extract PNG images, resize to 300x300, and add watermark.
        ProcessExtractedImages(docPath);
    }

    private static void CreateSamplePng(string filePath)
    {
        const int width = 200;
        const int height = 200;

        // Create bitmap and draw deterministic content.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5);
        graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
        bitmap.Save(filePath);

        // Cleanup.
        pen.Dispose();
        graphics.Dispose();
        bitmap.Dispose();
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    private static void ProcessExtractedImages(string docPath)
    {
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Save the extracted image to a temporary file.
            string extractedPath = $"extracted-{extractedCount}.png";
            shape.ImageData.Save(extractedPath);

            // Load the extracted image.
            using (Aspose.Drawing.Bitmap originalBitmap = new Aspose.Drawing.Bitmap(extractedPath))
            {
                // Create a new bitmap with the target size.
                Aspose.Drawing.Bitmap resizedBitmap = new Aspose.Drawing.Bitmap(300, 300);
                using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(resizedBitmap))
                {
                    // Draw the original image scaled to 300x300.
                    graphics.DrawImage(originalBitmap, 0, 0, 300, 300);

                    // Prepare watermark text.
                    string watermarkText = "Watermark";
                    Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20);
                    Aspose.Drawing.Brush brush = new Aspose.Drawing.SolidBrush(
                        Aspose.Drawing.Color.FromArgb(128, Aspose.Drawing.Color.Red));

                    // Draw watermark near the bottom-left corner.
                    graphics.DrawString(watermarkText, font, brush, new PointF(10, 260));
                }

                // Save the watermarked image.
                string watermarkedPath = $"watermarked-{extractedCount}.png";
                resizedBitmap.Save(watermarkedPath);
                resizedBitmap.Dispose();

                // Validate that the watermarked file was created.
                if (!File.Exists(watermarkedPath))
                    throw new Exception($"Failed to create watermarked image: {watermarkedPath}");
            }

            // Cleanup temporary extracted image.
            if (File.Exists(extractedPath))
                File.Delete(extractedPath);

            extractedCount++;
        }

        if (extractedCount == 0)
            throw new Exception("No images were extracted from the document.");
    }
}

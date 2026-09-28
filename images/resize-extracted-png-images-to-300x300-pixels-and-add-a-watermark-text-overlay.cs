using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample PNG image.
        string inputImagePath = "input.png";
        CreateSamplePng(inputImagePath);

        // Create a Word document and insert the sample image.
        string docPath = "sample.docx";
        CreateDocumentWithImage(docPath, inputImagePath);

        // Extract, resize to 300x300 and add a watermark overlay.
        ProcessDocumentImages(docPath);
    }

    private static void CreateSamplePng(string path)
    {
        int width = 500;
        int height = 400;

        // Use Aspose.Drawing types explicitly.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);

        graphics.Clear(Aspose.Drawing.Color.LightBlue);
        Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 48);
        graphics.DrawString("Sample", font, new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.DarkBlue), new Aspose.Drawing.PointF(50, 150));

        // Save the image and release resources.
        bitmap.Save(path);
        graphics.Dispose();
        bitmap.Dispose();

        if (!File.Exists(path))
            throw new Exception($"Failed to create sample image at {path}");
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        if (!File.Exists(docPath))
            throw new Exception($"Failed to create document at {docPath}");
    }

    private static void ProcessDocumentImages(string docPath)
    {
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            using (MemoryStream imageStream = new MemoryStream())
            {
                // Save the original image data to a stream.
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0;

                // Load the image into a bitmap.
                Aspose.Drawing.Bitmap originalBitmap = new Aspose.Drawing.Bitmap(imageStream);
                const int targetSize = 300;

                // Create a new bitmap for the resized image.
                Aspose.Drawing.Bitmap resizedBitmap = new Aspose.Drawing.Bitmap(targetSize, targetSize);
                Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(resizedBitmap);

                // Fill background and draw the resized original.
                graphics.Clear(Aspose.Drawing.Color.White);
                graphics.DrawImage(originalBitmap, new Aspose.Drawing.Rectangle(0, 0, targetSize, targetSize));

                // Add watermark text overlay.
                Aspose.Drawing.Font watermarkFont = new Aspose.Drawing.Font("Arial", 24);
                string watermark = "Watermark";
                Aspose.Drawing.SizeF textSize = graphics.MeasureString(watermark, watermarkFont);
                Aspose.Drawing.PointF position = new Aspose.Drawing.PointF(
                    (targetSize - textSize.Width) / 2,
                    (targetSize - textSize.Height) / 2);

                using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.FromArgb(128, Aspose.Drawing.Color.Red)))
                {
                    graphics.DrawString(watermark, watermarkFont, brush, position);
                }

                // Clean up drawing resources.
                graphics.Dispose();
                originalBitmap.Dispose();

                // Save the watermarked, resized image.
                string outputPath = $"output-{imageIndex}.png";
                resizedBitmap.Save(outputPath);
                resizedBitmap.Dispose();

                if (!File.Exists(outputPath))
                    throw new Exception($"Failed to save watermarked image at {outputPath}");
            }

            imageIndex++;
        }

        if (imageIndex == 0)
            throw new Exception("No images were extracted from the document.");
    }
}

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
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string docsDir = Path.Combine(baseDir, "Docs");
        string extractedDir = Path.Combine(baseDir, "Extracted");
        string watermarkedDir = Path.Combine(baseDir, "Watermarked");
        Directory.CreateDirectory(docsDir);
        Directory.CreateDirectory(extractedDir);
        Directory.CreateDirectory(watermarkedDir);

        // Create a deterministic sample PNG image.
        string sampleImagePath = Path.Combine(baseDir, "input.png");
        CreateSamplePng(sampleImagePath);

        // Create sample Word documents that contain the PNG image.
        CreateSampleDocument(Path.Combine(docsDir, "doc1.docx"), sampleImagePath);
        CreateSampleDocument(Path.Combine(docsDir, "doc2.docx"), sampleImagePath);

        // Extract PNG images from each document and apply watermark.
        int watermarkedCount = 0;
        foreach (string docPath in Directory.GetFiles(docsDir, "*.docx"))
        {
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

                // Save extracted image.
                string extractedImagePath = Path.Combine(
                    extractedDir,
                    $"{Path.GetFileNameWithoutExtension(docPath)}-image{imageIndex}.png");
                shape.ImageData.Save(extractedImagePath);
                if (!File.Exists(extractedImagePath))
                    throw new Exception($"Failed to extract image to {extractedImagePath}");

                // Apply watermark.
                string watermarkedImagePath = Path.Combine(
                    watermarkedDir,
                    $"{Path.GetFileNameWithoutExtension(docPath)}-image{imageIndex}-watermarked.png");
                ApplyWatermarkToPng(extractedImagePath, watermarkedImagePath, "WATERMARK");
                if (!File.Exists(watermarkedImagePath))
                    throw new Exception($"Failed to save watermarked image to {watermarkedImagePath}");

                watermarkedCount++;
                imageIndex++;
            }
        }

        // Validation.
        if (watermarkedCount == 0)
            throw new Exception("No watermarked images were produced.");
    }

    // Creates a deterministic sample PNG image using Aspose.Drawing.
    private static void CreateSamplePng(string filePath)
    {
        const int width = 200;
        const int height = 100;
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap);
        g.Clear(Aspose.Drawing.Color.White);
        using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.LightBlue))
        {
            g.FillRectangle(brush, 0, 0, width, height);
        }
        using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20))
        {
            using (Aspose.Drawing.SolidBrush textBrush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black))
            {
                g.DrawString("Sample", font, textBrush, new Aspose.Drawing.PointF(20, 30));
            }
        }
        bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        g.Dispose();
        bitmap.Dispose();

        if (!File.Exists(filePath))
            throw new Exception($"Failed to create sample image at {filePath}");
    }

    // Creates a Word document containing the specified image.
    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with an image:");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
        if (!File.Exists(docPath))
            throw new Exception($"Failed to save document at {docPath}");
    }

    // Loads a PNG, draws a watermark text, and saves the result.
    private static void ApplyWatermarkToPng(string inputPath, string outputPath, string watermarkText)
    {
        Aspose.Drawing.Bitmap original = new Aspose.Drawing.Bitmap(inputPath);
        Aspose.Drawing.Bitmap watermarked = new Aspose.Drawing.Bitmap(original.Width, original.Height);
        Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(watermarked);
        // Draw the original image.
        g.DrawImage(original, 0, 0, original.Width, original.Height);
        // Prepare watermark text.
        using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24, Aspose.Drawing.FontStyle.Bold))
        {
            // Semi‑transparent red brush.
            using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.FromArgb(128, 255, 0, 0)))
            {
                // Measure text size to position it at bottom‑right corner.
                Aspose.Drawing.SizeF textSize = g.MeasureString(watermarkText, font);
                float x = original.Width - textSize.Width - 10;
                float y = original.Height - textSize.Height - 10;
                g.DrawString(watermarkText, font, brush, new Aspose.Drawing.PointF(x, y));
            }
        }
        watermarked.Save(outputPath, Aspose.Drawing.Imaging.ImageFormat.Png);
        g.Dispose();
        original.Dispose();
        watermarked.Dispose();
    }
}

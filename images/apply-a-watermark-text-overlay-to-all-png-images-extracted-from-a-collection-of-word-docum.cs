using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a sample PNG image (used as source image in documents)
        string sampleImagePath = Path.Combine(outputDir, "sample.png");
        CreateSamplePng(sampleImagePath, 200, 100);

        // 2. Create sample Word documents that contain the PNG image
        string[] docPaths = new string[2];
        for (int i = 0; i < docPaths.Length; i++)
        {
            string docPath = Path.Combine(outputDir, $"Document{i + 1}.docx");
            CreateWordDocumentWithImage(docPath, sampleImagePath);
            docPaths[i] = docPath;
        }

        // 3. Process each document: extract PNG images, apply watermark, save result
        int docIndex = 0;
        foreach (string docPath in docPaths)
        {
            Document doc = new Document(docPath);
            var shapeNodes = doc.GetChildNodes(NodeType.Shape, true)
                                .Cast<Shape>()
                                .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Png)
                                .ToList();

            if (!shapeNodes.Any())
                throw new Exception($"No PNG images found in document '{docPath}'.");

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes)
            {
                // Extract image to memory stream
                using (MemoryStream imageStream = new MemoryStream())
                {
                    shape.ImageData.Save(imageStream);
                    imageStream.Position = 0; // reset before reading

                    // Load extracted image into a bitmap
                    using (Bitmap originalBitmap = new Bitmap(imageStream))
                    {
                        // Create a new bitmap to draw watermark onto
                        using (Bitmap watermarkedBitmap = new Bitmap(originalBitmap.Width, originalBitmap.Height))
                        {
                            using (Graphics graphics = Graphics.FromImage(watermarkedBitmap))
                            {
                                // Draw the original image
                                graphics.DrawImage(originalBitmap, 0, 0, originalBitmap.Width, originalBitmap.Height);

                                // Prepare watermark text drawing tools
                                string watermarkText = "Sample Watermark";
                                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20))
                                using (SolidBrush brush = new SolidBrush(Aspose.Drawing.Color.FromArgb(128, 255, 0, 0))) // semi‑transparent red
                                {
                                    // Position watermark near bottom‑left corner
                                    PointF position = new PointF(10, originalBitmap.Height - 30);
                                    graphics.DrawString(watermarkText, font, brush, position);
                                }
                            }

                            // Save the watermarked image
                            string watermarkedPath = Path.Combine(outputDir,
                                $"watermarked_doc{docIndex + 1}_img{imageIndex + 1}.png");
                            watermarkedBitmap.Save(watermarkedPath, ImageFormat.Png);

                            // Validate that the file was created
                            if (!File.Exists(watermarkedPath))
                                throw new Exception($"Failed to save watermarked image to '{watermarkedPath}'.");
                        }
                    }
                }

                imageIndex++;
            }

            docIndex++;
        }

        // Execution finished – all watermarked images are saved in the Output folder.
    }

    // Creates a deterministic PNG image file.
    private static void CreateSamplePng(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.White);
            using (SolidBrush brush = new SolidBrush(Aspose.Drawing.Color.LightBlue))
            {
                graphics.FillRectangle(brush, 0, 0, width, height);
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }

        if (!File.Exists(filePath))
            throw new Exception($"Failed to create sample image at '{filePath}'.");
    }

    // Creates a Word document containing the specified image.
    private static void CreateWordDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
        if (!File.Exists(docPath))
            throw new Exception($"Failed to save document at '{docPath}'.");
    }
}

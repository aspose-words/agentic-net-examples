using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string imagesDir = Path.Combine(baseDir, "Images");
        string docsDir = Path.Combine(baseDir, "Docs");
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(imagesDir);
        Directory.CreateDirectory(docsDir);
        Directory.CreateDirectory(outputDir);

        // 1. Create a sample PNG image (input.png)
        string sampleImagePath = Path.Combine(imagesDir, "input.png");
        CreateSamplePng(sampleImagePath, 300, 150, "Sample");

        // 2. Create sample Word documents that contain the PNG image
        int docCount = 2;
        for (int i = 0; i < docCount; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {i + 1}");
            // Insert the PNG image
            builder.InsertImage(sampleImagePath);
            string docPath = Path.Combine(docsDir, $"Doc{i + 1}.docx");
            doc.Save(docPath);
        }

        // 3. Process each document, extract PNG images, apply watermark, and save
        string[] docFiles = Directory.GetFiles(docsDir, "*.docx");
        int totalWatermarked = 0;

        for (int docIndex = 0; docIndex < docFiles.Length; docIndex++)
        {
            string docPath = docFiles[docIndex];
            Document doc = new Document(docPath);

            // Get all shape nodes that contain PNG images
            var shapeNodes = doc.GetChildNodes(NodeType.Shape, true)
                                .Cast<Shape>()
                                .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Png)
                                .ToList();

            for (int imgIndex = 0; imgIndex < shapeNodes.Count; imgIndex++)
            {
                Shape shape = shapeNodes[imgIndex];

                // Save the image to a memory stream
                using (MemoryStream imgStream = new MemoryStream())
                {
                    shape.ImageData.Save(imgStream);
                    imgStream.Position = 0; // Reset before reading

                    // Load the image into a bitmap
                    using (Aspose.Drawing.Bitmap originalBitmap = new Aspose.Drawing.Bitmap(imgStream))
                    {
                        // Create a new bitmap to avoid indexed pixel formats
                        using (Aspose.Drawing.Bitmap watermarkedBitmap = new Aspose.Drawing.Bitmap(originalBitmap.Width, originalBitmap.Height))
                        {
                            // Draw the original image onto the new bitmap
                            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(watermarkedBitmap))
                            {
                                graphics.DrawImage(originalBitmap, 0, 0, originalBitmap.Width, originalBitmap.Height);

                                // Prepare watermark text
                                string watermarkText = "CONFIDENTIAL";
                                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24, Aspose.Drawing.FontStyle.Bold))
                                using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.FromArgb(128, 255, 0, 0))) // Semi‑transparent red
                                {
                                    // Measure text size
                                    Aspose.Drawing.SizeF textSize = graphics.MeasureString(watermarkText, font);
                                    // Position text at the center
                                    float x = (watermarkedBitmap.Width - textSize.Width) / 2;
                                    float y = (watermarkedBitmap.Height - textSize.Height) / 2;
                                    graphics.DrawString(watermarkText, font, brush, x, y);
                                }
                            }

                            // Save the watermarked image
                            string watermarkedPath = Path.Combine(outputDir,
                                $"watermarked_doc{docIndex + 1}_img{imgIndex + 1}.png");
                            watermarkedBitmap.Save(watermarkedPath);

                            // Validate that the file was created
                            if (!File.Exists(watermarkedPath))
                                throw new InvalidOperationException($"Failed to create watermarked image: {watermarkedPath}");

                            totalWatermarked++;
                        }
                    }
                }
            }
        }

        // Ensure at least one watermarked image was produced
        if (totalWatermarked == 0)
            throw new InvalidOperationException("No PNG images were found and watermarked.");

        Console.WriteLine($"Processed {docFiles.Length} document(s) and created {totalWatermarked} watermarked PNG image(s).");
    }

    // Helper method to create a deterministic PNG image
    private static void CreateSamplePng(string filePath, int width, int height, string text)
    {
        // Create bitmap
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        // Create graphics object
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        // Fill background
        graphics.Clear(Aspose.Drawing.Color.White);

        // Draw centered text
        using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20, Aspose.Drawing.FontStyle.Regular))
        using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black))
        {
            Aspose.Drawing.SizeF textSize = graphics.MeasureString(text, font);
            float x = (width - textSize.Width) / 2;
            float y = (height - textSize.Height) / 2;
            graphics.DrawString(text, font, brush, x, y);
        }

        // Save to file
        bitmap.Save(filePath);

        // Clean up
        graphics.Dispose();
        bitmap.Dispose();
    }
}

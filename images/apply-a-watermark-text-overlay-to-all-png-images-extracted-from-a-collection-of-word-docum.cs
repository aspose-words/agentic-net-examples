using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string inputFolder = "InputDocs";
        string outputFolder = "WatermarkedImages";
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a sample PNG image (sample.png)
        string sampleImagePath = "sample.png";
        CreateSamplePng(sampleImagePath);

        // Create sample Word documents containing the PNG image
        CreateSampleDocument(Path.Combine(inputFolder, "doc1.docx"), sampleImagePath);
        CreateSampleDocument(Path.Combine(inputFolder, "doc2.docx"), sampleImagePath);

        // Process each document
        string[] docFiles = Directory.GetFiles(inputFolder, "*.docx");
        int docIndex = 0;
        List<string> watermarkedFiles = new List<string>();

        foreach (string docPath in docFiles)
        {
            Document doc = new Document(docPath);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapes)
            {
                if (!shape.HasImage)
                    continue;

                // Process only PNG images
                if (shape.ImageData.ImageType != ImageType.Png)
                    continue;

                // Extract the PNG image
                string extractedPath = Path.Combine(outputFolder,
                    $"extracted-{docIndex}-{imageIndex}.png");
                shape.ImageData.Save(extractedPath);

                // Apply watermark
                string watermarkedPath = Path.Combine(outputFolder,
                    $"watermarked-{docIndex}-{imageIndex}.png");
                ApplyWatermark(extractedPath, watermarkedPath);
                watermarkedFiles.Add(watermarkedPath);

                imageIndex++;
            }

            docIndex++;
        }

        // Validation: ensure at least one watermarked image was created
        if (watermarkedFiles.Count == 0)
            throw new InvalidOperationException("No watermarked images were produced.");

        Console.WriteLine($"Processed {docFiles.Length} documents, created {watermarkedFiles.Count} watermarked images.");
    }

    private static void CreateSamplePng(string filePath)
    {
        int width = 200;
        int height = 100;
        // Use Aspose.Drawing.Bitmap explicitly
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            // Create graphics from bitmap
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                // Fill background
                g.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.LightBlue))
                {
                    g.FillRectangle(brush, 0, 0, width, height);
                }
                // Draw sample text
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20))
                {
                    using (Aspose.Drawing.SolidBrush textBrush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.DarkBlue))
                    {
                        g.DrawString("Sample", font, textBrush, new Aspose.Drawing.PointF(20, 30));
                    }
                }
            }
            // Save as PNG
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        // Validate that the sample image was created
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create sample image: {filePath}");
    }

    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with image:");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    private static void ApplyWatermark(string sourcePath, string destPath)
    {
        // Load original image
        using (Aspose.Drawing.Bitmap original = new Aspose.Drawing.Bitmap(sourcePath))
        {
            // Create a new bitmap to draw onto (non-indexed)
            using (Aspose.Drawing.Bitmap watermarked = new Aspose.Drawing.Bitmap(original.Width, original.Height))
            {
                using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(watermarked))
                {
                    // Draw the original image
                    g.DrawImage(original, 0, 0, original.Width, original.Height);

                    // Prepare watermark text
                    string watermarkText = "Watermark";
                    using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
                    {
                        // Semi-transparent red brush
                        using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.FromArgb(128, Aspose.Drawing.Color.Red)))
                        {
                            // Position watermark at bottom-right corner
                            Aspose.Drawing.SizeF textSize = g.MeasureString(watermarkText, font);
                            float x = original.Width - textSize.Width - 10;
                            float y = original.Height - textSize.Height - 10;
                            g.DrawString(watermarkText, font, brush, new Aspose.Drawing.PointF(x, y));
                        }
                    }
                }

                // Save watermarked image
                watermarked.Save(destPath, Aspose.Drawing.Imaging.ImageFormat.Png);
            }
        }

        // Validate output file existence
        if (!File.Exists(destPath))
            throw new InvalidOperationException($"Failed to create watermarked image: {destPath}");
    }
}

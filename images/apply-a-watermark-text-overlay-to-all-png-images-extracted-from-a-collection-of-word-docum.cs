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
        string baseDir = Directory.GetCurrentDirectory();
        string docsDir = Path.Combine(baseDir, "InputDocs");
        string extractedDir = Path.Combine(baseDir, "ExtractedImages");
        string watermarkedDir = Path.Combine(baseDir, "WatermarkedImages");
        Directory.CreateDirectory(docsDir);
        Directory.CreateDirectory(extractedDir);
        Directory.CreateDirectory(watermarkedDir);

        // Create a sample PNG image to insert into documents
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSamplePng(sampleImagePath, 200, 100, "Sample");

        // Create sample Word documents containing the PNG image
        int docCount = 2;
        for (int i = 1; i <= docCount; i++)
        {
            string docPath = Path.Combine(docsDir, $"Document{i}.docx");
            CreateWordDocumentWithImage(docPath, sampleImagePath);
        }

        // Extract PNG images from all documents
        List<string> extractedPngPaths = new List<string>();
        foreach (string docFile in Directory.GetFiles(docsDir, "*.docx"))
        {
            Document doc = new Document(docFile);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage && shape.ImageData.ImageType == ImageType.Png)
                {
                    string extractedPath = Path.Combine(
                        extractedDir,
                        $"{Path.GetFileNameWithoutExtension(docFile)}_image{imageIndex}.png");
                    shape.ImageData.Save(extractedPath);
                    extractedPngPaths.Add(extractedPath);
                    imageIndex++;
                }
            }
        }

        if (extractedPngPaths.Count == 0)
            throw new Exception("No PNG images were extracted from the documents.");

        // Apply watermark to each extracted PNG image
        foreach (string pngPath in extractedPngPaths)
        {
            ApplyWatermark(pngPath, watermarkedDir, "WATERMARK");
        }

        // Validate that watermarked images were created
        string[] watermarkedFiles = Directory.GetFiles(watermarkedDir, "*.png");
        if (watermarkedFiles.Length == 0)
            throw new Exception("No watermarked images were created.");

        // Program completed
        Console.WriteLine($"Processed {docCount} documents.");
        Console.WriteLine($"Extracted {extractedPngPaths.Count} PNG images.");
        Console.WriteLine($"Created {watermarkedFiles.Length} watermarked images.");
    }

    private static void CreateSamplePng(string path, int width, int height, string text)
    {
        // Create bitmap using Aspose.Drawing
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20))
                {
                    using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black))
                    {
                        g.DrawString(text, font, brush, new Aspose.Drawing.PointF(10, 10));
                    }
                }
            }
            bitmap.Save(path, Aspose.Drawing.Imaging.ImageFormat.Png);
        }
    }

    private static void CreateWordDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document containing an image:");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    private static void ApplyWatermark(string sourceImagePath, string outputDir, string watermarkText)
    {
        // Load original image
        using (Aspose.Drawing.Bitmap original = (Aspose.Drawing.Bitmap)Aspose.Drawing.Image.FromFile(sourceImagePath))
        {
            // Create a new bitmap with same size and 32bpp ARGB pixel format
            using (Aspose.Drawing.Bitmap watermarked = new Aspose.Drawing.Bitmap(
                original.Width, original.Height, Aspose.Drawing.Imaging.PixelFormat.Format32bppArgb))
            {
                using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(watermarked))
                {
                    // Draw the original image onto the new bitmap
                    g.DrawImage(original, 0, 0, original.Width, original.Height);

                    // Prepare watermark drawing tools
                    using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24, Aspose.Drawing.FontStyle.Bold))
                    {
                        using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(
                            Aspose.Drawing.Color.FromArgb(128, Aspose.Drawing.Color.LightGray)))
                        {
                            // Measure text size
                            Aspose.Drawing.SizeF textSize = g.MeasureString(watermarkText, font);
                            // Position at bottom-right with some padding
                            float x = watermarked.Width - textSize.Width - 10;
                            float y = watermarked.Height - textSize.Height - 10;
                            g.DrawString(watermarkText, font, brush, new Aspose.Drawing.PointF(x, y));
                        }
                    }
                }

                // Save watermarked image
                string fileName = Path.GetFileName(sourceImagePath);
                string outputPath = Path.Combine(outputDir, $"watermarked_{fileName}");
                watermarked.Save(outputPath, Aspose.Drawing.Imaging.ImageFormat.Png);
            }
        }
    }
}

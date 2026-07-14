using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare output folder
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(baseDir);

        // 1. Create a sample PNG image (input.png)
        string sampleImagePath = Path.Combine(baseDir, "input.png");
        CreateSamplePng(sampleImagePath, 200, 100);

        // 2. Create sample Word documents that contain the PNG image
        List<string> docPaths = new List<string>();
        for (int i = 0; i < 2; i++)
        {
            string docPath = Path.Combine(baseDir, $"Document{i + 1}.docx");
            CreateWordDocumentWithImage(docPath, sampleImagePath);
            docPaths.Add(docPath);
        }

        // 3. Process each document: extract PNG images, apply watermark, save result
        int docIndex = 0;
        foreach (string docPath in docPaths)
        {
            Document doc = new Document(docPath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage && shape.ImageData.ImageType == ImageType.Png)
                {
                    // Extract the image to a memory stream
                    using (MemoryStream imageStream = new MemoryStream())
                    {
                        shape.ImageData.Save(imageStream);
                        imageStream.Position = 0; // Reset before reading

                        // Load the extracted image into a bitmap
                        using (Aspose.Drawing.Bitmap originalBitmap = new Aspose.Drawing.Bitmap(imageStream))
                        {
                            // Create a new bitmap to draw the watermark onto
                            using (Aspose.Drawing.Bitmap watermarkedBitmap = new Aspose.Drawing.Bitmap(originalBitmap.Width, originalBitmap.Height))
                            {
                                using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(watermarkedBitmap))
                                {
                                    // Draw the original image onto the new bitmap
                                    graphics.DrawImage(originalBitmap, 0, 0, originalBitmap.Width, originalBitmap.Height);

                                    // Prepare watermark text drawing
                                    string watermarkText = "Sample Watermark";
                                    Aspose.Drawing.Font watermarkFont = new Aspose.Drawing.Font("Arial", 20);
                                    Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.FromArgb(128, Aspose.Drawing.Color.Red));
                                    Aspose.Drawing.PointF position = new Aspose.Drawing.PointF(10, 10);

                                    // Draw the watermark text
                                    graphics.DrawString(watermarkText, watermarkFont, brush, position);
                                }

                                // Save the watermarked image
                                string watermarkedPath = Path.Combine(
                                    baseDir,
                                    $"watermarked_doc{docIndex + 1}_img{imageIndex + 1}.png");
                                watermarkedBitmap.Save(watermarkedPath, Aspose.Drawing.Imaging.ImageFormat.Png);

                                // Validate that the file was created
                                if (!File.Exists(watermarkedPath))
                                    throw new Exception($"Failed to create watermarked image: {watermarkedPath}");
                            }
                        }
                    }

                    imageIndex++;
                }
            }

            docIndex++;
        }
    }

    // Creates a simple white PNG image with given dimensions
    private static void CreateSamplePng(string filePath, int width, int height)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.White);
            }
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        if (!File.Exists(filePath))
            throw new Exception($"Failed to create sample image: {filePath}");
    }

    // Creates a Word document containing the specified image
    private static void CreateWordDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
        if (!File.Exists(docPath))
            throw new Exception($"Failed to create document: {docPath}");
    }
}

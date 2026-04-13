using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Directory for all generated files.
        string dataDir = "Data";
        Directory.CreateDirectory(dataDir);

        // -----------------------------------------------------------------
        // 1. Create a sample PNG image that will be inserted into documents.
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(dataDir, "sample.png");
        CreateSamplePng(sampleImagePath);

        // -----------------------------------------------------------------
        // 2. Create a sample Word document containing the PNG image.
        // -----------------------------------------------------------------
        string docPath = Path.Combine(dataDir, "sample.docx");
        CreateSampleDocument(docPath, sampleImagePath);

        // -----------------------------------------------------------------
        // 3. Process all Word documents in the directory:
        //    - Extract PNG images.
        //    - Apply a text watermark overlay.
        //    - Save the watermarked images.
        // -----------------------------------------------------------------
        int watermarkedCount = 0;
        foreach (string file in Directory.GetFiles(dataDir, "*.docx"))
        {
            Document doc = new Document(file);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage && shape.ImageData.ImageType == ImageType.Png)
                {
                    // Extract the image to a memory stream.
                    using (MemoryStream imageStream = new MemoryStream())
                    {
                        shape.ImageData.Save(imageStream);
                        imageStream.Position = 0; // Reset before reading.

                        // Load the extracted image into a bitmap.
                        using (Aspose.Drawing.Bitmap originalBitmap = new Aspose.Drawing.Bitmap(imageStream))
                        {
                            // Create a new bitmap to draw the watermark onto.
                            using (Aspose.Drawing.Bitmap watermarkedBitmap = new Aspose.Drawing.Bitmap(originalBitmap.Width, originalBitmap.Height))
                            {
                                using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(watermarkedBitmap))
                                {
                                    // Draw the original image.
                                    graphics.DrawImage(originalBitmap, 0, 0, originalBitmap.Width, originalBitmap.Height);

                                    // Prepare watermark text drawing.
                                    using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20))
                                    using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.FromArgb(128, Aspose.Drawing.Color.Red)))
                                    {
                                        // Position the watermark near the top‑left corner.
                                        graphics.DrawString("WATERMARK", font, brush, new Aspose.Drawing.PointF(10, 10));
                                    }
                                }

                                // Save the watermarked image.
                                string watermarkedPath = Path.Combine(dataDir, $"watermarked_{Path.GetFileNameWithoutExtension(file)}_{imageIndex}.png");
                                watermarkedBitmap.Save(watermarkedPath);
                                if (!File.Exists(watermarkedPath))
                                    throw new Exception($"Failed to save watermarked image: {watermarkedPath}");

                                watermarkedCount++;
                                imageIndex++;
                            }
                        }
                    }
                }
            }
        }

        // Validate that at least one watermarked image was produced.
        if (watermarkedCount == 0)
            throw new Exception("No PNG images were found and watermarked.");
    }

    // Creates a deterministic PNG image file.
    private static void CreateSamplePng(string filePath)
    {
        int width = 200;
        int height = 100;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.White);
            using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 3))
            {
                graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
            }
            bitmap.Save(filePath);
        }

        if (!File.Exists(filePath))
            throw new Exception($"Failed to create sample image: {filePath}");
    }

    // Creates a Word document that contains the specified image.
    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the PNG image twice to demonstrate multiple occurrences.
        builder.InsertImage(imagePath);
        builder.InsertParagraph();
        builder.InsertImage(imagePath);

        doc.Save(docPath);
        if (!File.Exists(docPath))
            throw new Exception($"Failed to create sample document: {docPath}");
    }
}

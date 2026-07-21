using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Create a sample PNG image.
        const string sampleImagePath = "sample.png";
        CreateSamplePng(sampleImagePath);

        // Create a Word document and insert the sample image.
        const string docPath = "DocumentWithImages.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // Extract PNG images, apply a red 5‑pixel border, and save them.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Load the original image bytes.
            byte[] originalBytes = shape.ImageData.ImageBytes;
            using (MemoryStream originalStream = new MemoryStream(originalBytes))
            {
                originalStream.Position = 0;
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    int borderSize = 5;
                    int newWidth = originalBitmap.Width + borderSize * 2;
                    int newHeight = originalBitmap.Height + borderSize * 2;

                    using (Bitmap borderedBitmap = new Bitmap(newWidth, newHeight))
                    {
                        using (Graphics g = Graphics.FromImage(borderedBitmap))
                        {
                            // Fill background with red (the border).
                            g.Clear(Color.Red);

                            // Draw the original image onto the new bitmap, offset by the border size.
                            g.DrawImage(originalBitmap, borderSize, borderSize, originalBitmap.Width, originalBitmap.Height);
                        }

                        // Save the bordered image.
                        string outputPath = $"extracted_{extractedCount + 1}.png";
                        borderedBitmap.Save(outputPath, ImageFormat.Png);
                        extractedCount++;
                    }
                }
            }
        }

        // Validate that at least one image was processed.
        if (extractedCount == 0)
            throw new InvalidOperationException("No PNG images were extracted and processed.");

        // Optional cleanup.
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }

    private static void CreateSamplePng(string path)
    {
        int width = 100;
        int height = 100;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // White background.
                g.Clear(Color.White);
                // Draw a simple black rectangle.
                using (Pen pen = new Pen(Color.Black, 2))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(path, ImageFormat.Png);
        }
    }
}

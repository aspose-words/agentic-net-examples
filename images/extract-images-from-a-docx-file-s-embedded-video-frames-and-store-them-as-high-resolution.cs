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
        // Create a sample high‑resolution image that will act as a video frame.
        const string frameImagePath = "frame.png";
        const int width = 800;
        const int height = 600;

        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);

                // Draw a simple rectangle with some text to identify the frame.
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 50, 50, width - 100, height - 100);
                }

                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 48, Aspose.Drawing.FontStyle.Bold))
                {
                    using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.DarkRed))
                    {
                        g.DrawString("Video Frame", font, brush, new Aspose.Drawing.PointF(150, height / 2 - 30));
                    }
                }
            }

            bitmap.Save(frameImagePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        // Create a DOCX document and insert the sample image (simulating a video frame).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(frameImagePath);
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // Reload the document to simulate a separate extraction step.
        Document loadedDoc = new Document(docPath);

        // Extract all images from the document (including those representing video frames).
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                string outputImagePath = $"extracted-{extractedCount + 1}.png";
                shape.ImageData.Save(outputImagePath);
                if (!File.Exists(outputImagePath))
                {
                    throw new InvalidOperationException($"Failed to save extracted image to '{outputImagePath}'.");
                }
                extractedCount++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
        {
            throw new InvalidOperationException("No images were extracted from the document.");
        }

        // Optional cleanup (commented out to keep output files for verification).
        // File.Delete(frameImagePath);
        // File.Delete(docPath);
    }
}

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
        const string inputImagePath = "input.png";
        using (Aspose.Drawing.Bitmap sampleBitmap = new Aspose.Drawing.Bitmap(200, 200))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(sampleBitmap))
            {
                g.Clear(Aspose.Drawing.Color.LightBlue);
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.DarkBlue, 5))
                {
                    g.DrawRectangle(pen, 20, 20, 160, 160);
                }
            }
            sampleBitmap.Save(inputImagePath);
        }

        // Insert the sample image into a Word document.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // Load the document and extract images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Save the shape's image to a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0;

                // Load the extracted image into a bitmap.
                using (Aspose.Drawing.Bitmap originalBitmap = new Aspose.Drawing.Bitmap(imageStream))
                {
                    // Create a new 300x300 bitmap for resizing.
                    using (Aspose.Drawing.Bitmap resizedBitmap = new Aspose.Drawing.Bitmap(300, 300))
                    {
                        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(resizedBitmap))
                        {
                            // Draw the original image scaled to 300x300.
                            graphics.DrawImage(originalBitmap, new Aspose.Drawing.Rectangle(0, 0, 300, 300));

                            // Add watermark text overlay.
                            using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
                            using (Aspose.Drawing.Brush brush = new Aspose.Drawing.SolidBrush(
                                Aspose.Drawing.Color.FromArgb(128, Aspose.Drawing.Color.White)))
                            {
                                graphics.DrawString("Watermark", font, brush, new Aspose.Drawing.PointF(10, 260));
                            }
                        }

                        // Save the watermarked, resized image.
                        string outputImagePath = $"output_{imageIndex}.png";
                        resizedBitmap.Save(outputImagePath);

                        // Validate that the output file was created.
                        if (!File.Exists(outputImagePath))
                            throw new InvalidOperationException($"Failed to create output image: {outputImagePath}");
                    }
                }
            }

            imageIndex++;
        }
    }
}

using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample PNG image (500x500) to be used as input.
        const string inputImagePath = "input.png";
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(500, 500))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 48))
                {
                    g.DrawString("Sample", font, new SolidBrush(Aspose.Drawing.Color.Black), new PointF(50, 200));
                }
            }
            bitmap.Save(inputImagePath);
        }

        // Insert the sample image into a new Word document.
        const string docPath = "doc_with_image.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // Load the document and extract PNG images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;
            if (shape.ImageData.ImageType != ImageType.Png) continue;

            // Save the extracted PNG image.
            string extractedPath = $"extracted_{imageIndex}.png";
            shape.ImageData.Save(extractedPath);

            // Load the extracted image, resize to 300x300 and add a watermark.
            using (Aspose.Drawing.Bitmap original = new Aspose.Drawing.Bitmap(extractedPath))
            {
                using (Aspose.Drawing.Bitmap resized = new Aspose.Drawing.Bitmap(300, 300))
                {
                    using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(resized))
                    {
                        // Fill background.
                        graphics.Clear(Aspose.Drawing.Color.White);
                        // Draw the original image scaled to 300x300.
                        graphics.DrawImage(original, new Rectangle(0, 0, 300, 300));

                        // Add watermark text overlay.
                        using (Aspose.Drawing.Font watermarkFont = new Aspose.Drawing.Font("Arial", 24))
                        {
                            // Semi‑transparent white text.
                            using (SolidBrush brush = new SolidBrush(Aspose.Drawing.Color.FromArgb(128, Aspose.Drawing.Color.White)))
                            {
                                graphics.DrawString("Watermark", watermarkFont, brush, new PointF(10, 260));
                            }
                        }
                    }

                    // Save the watermarked, resized image.
                    string watermarkedPath = $"watermarked_{imageIndex}.png";
                    resized.Save(watermarkedPath);

                    // Replace the image in the document with the new watermarked version.
                    shape.ImageData.SetImage(watermarkedPath);
                }
            }

            imageIndex++;
        }

        // Save the updated document.
        const string updatedDocPath = "doc_with_image_updated.docx";
        loadedDoc.Save(updatedDocPath);
    }
}

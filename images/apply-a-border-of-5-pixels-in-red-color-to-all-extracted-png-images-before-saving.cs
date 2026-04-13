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
        // Create a deterministic PNG image.
        const int imgWidth = 200;
        const int imgHeight = 200;
        string inputImagePath = "input.png";

        Bitmap bitmap = new Bitmap(imgWidth, imgHeight);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        // Draw a simple blue ellipse for visual content.
        using (Pen pen = new Pen(Color.Blue, 3))
        {
            graphics.DrawEllipse(pen, 20, 20, imgWidth - 40, imgHeight - 40);
        }
        bitmap.Save(inputImagePath);
        graphics.Dispose();
        bitmap.Dispose();

        // Create a document and insert the PNG image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        string docPath = "Document.docx";
        doc.Save(docPath);

        // Reload the document (optional, demonstrates load lifecycle).
        Document loadedDoc = new Document(docPath);

        // Extract PNG images, apply a 5‑pixel red border, and save them.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Save the original image to a memory stream.
            using (MemoryStream ms = new MemoryStream())
            {
                shape.ImageData.Save(ms);
                ms.Position = 0;

                // Load the image with Aspose.Drawing for manipulation.
                using (Image img = Image.FromStream(ms))
                {
                    using (Graphics g = Graphics.FromImage(img))
                    {
                        // Draw a red rectangle border of 5 pixels.
                        using (Pen borderPen = new Pen(Color.Red, 5))
                        {
                            // Adjust rectangle to stay within image bounds.
                            g.DrawRectangle(borderPen, 0, 0, img.Width - 1, img.Height - 1);
                        }
                    }

                    // Save the modified image to a deterministic file name.
                    string outFileName = $"extracted_{extractedCount}.png";
                    img.Save(outFileName, ImageFormat.Png);
                    extractedCount++;
                }
            }
        }

        // Validation: ensure at least one image was extracted and saved.
        if (extractedCount == 0)
            throw new InvalidOperationException("No PNG images were extracted from the document.");

        // Cleanup: optional deletion of temporary files can be added here.
    }
}

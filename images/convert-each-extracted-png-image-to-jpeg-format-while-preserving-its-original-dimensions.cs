using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a sample PNG image.
        string pngPath = "sample.png";
        using (var bitmap = new Aspose.Drawing.Bitmap(200, 100))
        {
            using (var graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.LightBlue);
                var font = new Aspose.Drawing.Font("Arial", 20);
                var brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black);
                graphics.DrawString("Sample", font, brush, new Aspose.Drawing.PointF(10, 40));
            }
            bitmap.Save(pngPath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        // Insert the PNG image into a new document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(pngPath);
        string docPath = "DocumentWithPng.docx";
        doc.Save(docPath);

        // Load the document and extract PNG images.
        var loadedDoc = new Document(docPath);
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Png)
            {
                using (var ms = new MemoryStream())
                {
                    shape.ImageData.Save(ms);
                    ms.Position = 0; // Reset stream before reading.

                    using (var bitmap = new Aspose.Drawing.Bitmap(ms))
                    {
                        string jpegPath = $"extracted_{imageIndex}.jpg";
                        bitmap.Save(jpegPath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
                        imageIndex++;
                    }
                }
            }
        }

        if (imageIndex == 0)
            throw new InvalidOperationException("No PNG images were extracted.");
    }
}

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
        // Create a sample JPEG image.
        const string jpegPath = "sample.jpg";
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                using (Pen pen = new Pen(Color.Red, 5))
                {
                    graphics.DrawEllipse(pen, 20, 20, 160, 160);
                }
            }
            bitmap.Save(jpegPath, ImageFormat.Jpeg);
        }

        // Create a Word document and insert the JPEG image.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(jpegPath);
        doc.Save(docPath);

        // Load the document and extract images, converting each JPEG to WebP.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int convertedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                byte[] imageBytes = shape.ImageData.ImageBytes;
                using (MemoryStream ms = new MemoryStream(imageBytes))
                {
                    ms.Position = 0;
                    using (Image img = Image.FromStream(ms))
                    {
                        string webpPath = $"converted-{convertedCount}.webp";
                        img.Save(webpPath); // Save inferred as WebP by extension.
                        if (!File.Exists(webpPath))
                        {
                            throw new Exception($"Failed to save WebP image: {webpPath}");
                        }
                        convertedCount++;
                    }
                }
            }
        }

        if (convertedCount == 0)
        {
            throw new Exception("No images were extracted and converted.");
        }
    }
}

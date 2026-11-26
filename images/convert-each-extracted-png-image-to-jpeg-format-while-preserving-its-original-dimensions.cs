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
        // Create sample PNG images.
        CreateSamplePng("image1.png", 200, 150, Aspose.Drawing.Color.LightBlue);
        CreateSamplePng("image2.png", 300, 100, Aspose.Drawing.Color.LightGreen);

        // Create a new document and insert the PNG images.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage("image1.png");
        builder.Writeln(); // separate images
        builder.InsertImage("image2.png");
        doc.Save("DocumentWithImages.docx");

        // Load the document (optional, we already have it).
        Document loadedDoc = new Document("DocumentWithImages.docx");

        // Extract PNG images and convert each to JPEG.
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int jpegIndex = 1;
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Png)
            {
                // Get PNG image bytes.
                byte[] pngBytes = shape.ImageData.ImageBytes;

                // Load PNG into Aspose.Drawing.Bitmap.
                using (MemoryStream ms = new MemoryStream(pngBytes))
                {
                    ms.Position = 0;
                    using (Bitmap bitmap = new Bitmap(ms))
                    {
                        // Save as JPEG preserving dimensions.
                        string jpegPath = $"extracted-{jpegIndex}.jpg";
                        bitmap.Save(jpegPath, ImageFormat.Jpeg);
                        jpegIndex++;
                    }
                }
            }
        }

        // Validate that at least one JPEG was created.
        if (jpegIndex == 1)
            throw new InvalidOperationException("No PNG images were found to convert.");

        // Optional: indicate completion.
        Console.WriteLine("PNG to JPEG conversion completed.");
    }

    private static void CreateSamplePng(string filePath, int width, int height, Aspose.Drawing.Color backgroundColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(backgroundColor);
                // Draw a simple rectangle for visual distinction.
                using (Pen pen = new Pen(Aspose.Drawing.Color.Black, 3))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}

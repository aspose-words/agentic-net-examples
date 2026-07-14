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
        // Create a deterministic PNG image using Aspose.Drawing.
        const string pngPath = "sample.png";
        CreateSamplePng(pngPath);

        // Insert the PNG image into a Word document.
        const string docPath = "document.docx";
        InsertImageIntoDocument(pngPath, docPath);

        // Load the document and convert each extracted PNG image to JPEG.
        ConvertExtractedPngsToJpeg(docPath);
    }

    private static void CreateSamplePng(string filePath)
    {
        // 200x200 white bitmap with a red ellipse.
        var bitmap = new Bitmap(200, 200);
        var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        using (var pen = new Pen(Color.Red, 5))
        {
            graphics.DrawEllipse(pen, 10, 10, 180, 180);
        }
        graphics.Dispose();
        bitmap.Save(filePath, ImageFormat.Png);
        bitmap.Dispose();
    }

    private static void InsertImageIntoDocument(string imagePath, string docPath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    private static void ConvertExtractedPngsToJpeg(string docPath)
    {
        var doc = new Document(docPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true)
                        .OfType<Shape>()
                        .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Png)
                        .ToList();

        int imageIndex = 0;
        foreach (var shape in shapes)
        {
            using (var ms = new MemoryStream())
            {
                // Save the original PNG bytes to a memory stream.
                shape.ImageData.Save(ms);
                ms.Position = 0;

                // Load the image via Aspose.Drawing and save as JPEG.
                using (var img = Image.FromStream(ms))
                {
                    string jpegPath = $"extracted_{imageIndex}.jpg";
                    img.Save(jpegPath, ImageFormat.Jpeg);
                    imageIndex++;
                }
            }
        }

        if (imageIndex == 0)
            throw new InvalidOperationException("No PNG images were found to convert.");
    }
}

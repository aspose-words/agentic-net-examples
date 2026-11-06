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
        // Create a sample PNG image.
        const string pngPath = "sample.png";
        CreateSamplePng(pngPath);

        // Create a Word document and insert the PNG image.
        const string docPath = "DocumentWithImages.docx";
        CreateDocumentWithImage(docPath, pngPath);

        // Load the document, extract PNG images, and convert each to JPEG.
        ConvertPngImagesToJpeg(docPath);
    }

    private static void CreateSamplePng(string path)
    {
        const int width = 200;
        const int height = 100;

        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                using (Pen pen = new Pen(Color.Blue, 3))
                {
                    graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }

            bitmap.Save(path, ImageFormat.Png);
        }

        if (!File.Exists(path))
            throw new Exception($"Failed to create PNG image at '{path}'.");
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        if (!File.Exists(docPath))
            throw new Exception($"Failed to save document at '{docPath}'.");
    }

    private static void ConvertPngImagesToJpeg(string docPath)
    {
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            byte[] imageBytes = shape.ImageData.ImageBytes;

            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0;
                using (Bitmap bitmap = new Bitmap(ms))
                {
                    string jpegPath = $"extracted-{imageIndex}.jpg";
                    bitmap.Save(jpegPath, ImageFormat.Jpeg);

                    if (!File.Exists(jpegPath))
                        throw new Exception($"Failed to save JPEG image at '{jpegPath}'.");
                }
            }

            imageIndex++;
        }

        if (imageIndex == 0)
            throw new Exception("No PNG images were found for conversion.");
    }
}

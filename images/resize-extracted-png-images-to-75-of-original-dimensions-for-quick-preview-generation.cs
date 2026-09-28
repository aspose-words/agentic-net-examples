using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a sample PNG image.
        const string inputImagePath = "input.png";
        CreateSamplePng(inputImagePath);

        // Create a Word document and insert the sample image.
        const string docPath = "document.docx";
        CreateDocumentWithImage(docPath, inputImagePath);

        // Extract PNG images, resize them to 75% of original size, and save previews.
        ExtractAndResizeImages(docPath);
    }

    private static void CreateSamplePng(string path)
    {
        const int width = 200;
        const int height = 200;

        using (var bitmap = new Bitmap(width, height))
        {
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                using (var pen = new Pen(Color.Blue, 5))
                {
                    graphics.DrawEllipse(pen, 10, 10, width - 20, height - 20);
                }
            }

            bitmap.Save(path);
        }
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    private static void ExtractAndResizeImages(string docPath)
    {
        var doc = new Document(docPath);
        var shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            using (var originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                using (var originalBitmap = new Bitmap(originalStream))
                {
                    int newWidth = (int)(originalBitmap.Width * 0.75);
                    int newHeight = (int)(originalBitmap.Height * 0.75);

                    using (var resizedBitmap = new Bitmap(newWidth, newHeight))
                    {
                        using (var graphics = Graphics.FromImage(resizedBitmap))
                        {
                            graphics.Clear(Color.Transparent);
                            graphics.DrawImage(originalBitmap, 0, 0, newWidth, newHeight);
                        }

                        string previewPath = $"preview-{imageIndex}.png";
                        resizedBitmap.Save(previewPath);

                        if (!File.Exists(previewPath))
                            throw new Exception($"Failed to save preview image: {previewPath}");
                    }
                }
            }

            imageIndex++;
        }

        if (imageIndex == 0)
            throw new Exception("No images were extracted from the document.");
    }
}

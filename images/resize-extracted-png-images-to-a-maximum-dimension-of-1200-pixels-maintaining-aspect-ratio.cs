using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Create a sample PNG image larger than the target size.
        string inputImagePath = "input.png";
        CreateSampleImage(inputImagePath, 2000, 1500);

        // Create a Word document and insert the sample image.
        string docPath = "sample.docx";
        CreateDocumentWithImage(docPath, inputImagePath);

        // Extract PNG images from the document, resize them, and save the results.
        string outputFolder = "output";
        ResizeExtractedPngImages(docPath, outputFolder);
    }

    private static void CreateSampleImage(string path, int width, int height)
    {
        var bitmap = new Bitmap(width, height);
        var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.LightBlue);
        graphics.Dispose();

        bitmap.Save(path, ImageFormat.Png);
        bitmap.Dispose();

        if (!File.Exists(path))
            throw new Exception($"Failed to create sample image at '{path}'.");
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        if (!File.Exists(docPath))
            throw new Exception($"Failed to save document at '{docPath}'.");
    }

    private static void ResizeExtractedPngImages(string docPath, string outputFolder)
    {
        if (!Directory.Exists(outputFolder))
            Directory.CreateDirectory(outputFolder);

        var doc = new Document(docPath);
        var shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            using (var imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0;

                using (var originalBitmap = new Bitmap(imageStream))
                {
                    int originalWidth = originalBitmap.Width;
                    int originalHeight = originalBitmap.Height;
                    const int maxDimension = 1200;
                    double scale = 1.0;

                    if (originalWidth > maxDimension || originalHeight > maxDimension)
                    {
                        scale = Math.Min((double)maxDimension / originalWidth, (double)maxDimension / originalHeight);
                    }

                    int newWidth = (int)(originalWidth * scale);
                    int newHeight = (int)(originalHeight * scale);

                    Bitmap resizedBitmap;

                    if (scale < 1.0)
                    {
                        resizedBitmap = new Bitmap(newWidth, newHeight);
                        using (var g = Graphics.FromImage(resizedBitmap))
                        {
                            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            g.DrawImage(originalBitmap, 0, 0, newWidth, newHeight);
                        }
                        originalBitmap.Dispose();
                    }
                    else
                    {
                        resizedBitmap = new Bitmap(originalBitmap);
                    }

                    string outputPath = Path.Combine(outputFolder, $"image-{imageIndex}.png");
                    resizedBitmap.Save(outputPath, ImageFormat.Png);
                    resizedBitmap.Dispose();

                    if (!File.Exists(outputPath))
                        throw new Exception($"Resized image was not saved at '{outputPath}'.");

                    imageIndex++;
                }
            }
        }

        if (imageIndex == 0)
            throw new Exception("No images were extracted from the document.");
    }
}

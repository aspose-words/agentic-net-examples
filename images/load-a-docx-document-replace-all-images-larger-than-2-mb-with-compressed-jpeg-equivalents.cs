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
        // Create deterministic sample images
        const string largeImagePath = "large.png";
        const string smallImagePath = "small.png";

        CreateSampleImage(3000, 3000, largeImagePath, true);   // noisy image to increase file size
        CreateSampleImage(100, 100, smallImagePath, false);   // simple small image

        // Create a sample DOCX containing both images
        const string inputDocPath = "input.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Document with images:");
        builder.InsertImage(largeImagePath);
        builder.Writeln();
        builder.InsertImage(smallImagePath);
        doc.Save(inputDocPath);

        // Load the document for processing
        var loadedDoc = new Document(inputDocPath);
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int replacedCount = 0;

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage) continue;

            byte[] imageBytes = shape.ImageData.ImageBytes;
            if (imageBytes.Length <= 2 * 1024 * 1024) continue; // Skip images <= 2 MB

            // Load the original image into a bitmap
            using (var originalStream = new MemoryStream(imageBytes))
            {
                originalStream.Position = 0;
                using (var bitmap = new Bitmap(originalStream))
                {
                    // Save the bitmap as a JPEG to a new memory stream
                    using (var jpegStream = new MemoryStream())
                    {
                        bitmap.Save(jpegStream, ImageFormat.Jpeg);
                        jpegStream.Position = 0;

                        // Replace the shape's image with the compressed JPEG
                        shape.ImageData.SetImage(jpegStream);
                        replacedCount++;
                    }
                }
            }
        }

        // Validate that at least one image was replaced
        if (replacedCount == 0)
            throw new InvalidOperationException("No images larger than 2 MB were found to replace.");

        // Save the modified document
        const string outputDocPath = "output.docx";
        loadedDoc.Save(outputDocPath);

        // Validate output file existence
        if (!File.Exists(outputDocPath))
            throw new FileNotFoundException("The output document was not created.", outputDocPath);
    }

    private static void CreateSampleImage(int width, int height, string filePath, bool noisy)
    {
        using (var bitmap = new Bitmap(width, height))
        {
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.LightBlue);
                if (noisy)
                {
                    var rnd = new Random();
                    for (int y = 0; y < height; y++)
                    {
                        for (int x = 0; x < width; x++)
                        {
                            var r = (byte)rnd.Next(256);
                            var g = (byte)rnd.Next(256);
                            var b = (byte)rnd.Next(256);
                            bitmap.SetPixel(x, y, Aspose.Drawing.Color.FromArgb(r, g, b));
                        }
                    }
                }
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}

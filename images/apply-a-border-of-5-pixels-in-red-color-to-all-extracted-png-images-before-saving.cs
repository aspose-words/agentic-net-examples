using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample PNG image.
        const int imgWidth = 200;
        const int imgHeight = 100;
        var sampleBitmap = new Bitmap(imgWidth, imgHeight);
        var sampleGraphics = Graphics.FromImage(sampleBitmap);
        sampleGraphics.Clear(Color.White);
        sampleBitmap.Save("input.png", ImageFormat.Png);
        sampleGraphics.Dispose();
        sampleBitmap.Dispose();

        // Create a Word document and insert the sample image.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertImage("input.png");
        doc.Save("DocumentWithImage.docx");

        // Load the document and process all PNG images.
        var loadedDoc = new Document("DocumentWithImage.docx");
        var shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int processedCount = 0;
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Extract the image to a memory stream.
            using (var imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0;

                // Load the image into a bitmap for manipulation.
                using (var bitmap = new Bitmap(imageStream))
                {
                    // Draw a 5‑pixel red border around the image.
                    using (var graphics = Graphics.FromImage(bitmap))
                    using (var pen = new Pen(Color.Red, 5))
                    {
                        graphics.DrawRectangle(pen, 0, 0, bitmap.Width - 1, bitmap.Height - 1);
                    }

                    // Save the modified image.
                    string outputPath = $"extracted_{imageIndex}.png";
                    bitmap.Save(outputPath, ImageFormat.Png);
                    processedCount++;
                }
            }

            imageIndex++;
        }

        // Validate that at least one image was processed.
        if (processedCount == 0)
            throw new Exception("No PNG images were extracted and processed.");
    }
}

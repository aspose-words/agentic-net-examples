using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a sample BMP image (800x600) and save it locally.
        const string inputBmpPath = "input.bmp";
        using (var bitmap = new Bitmap(800, 600))
        {
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
                using (var pen = new Pen(Color.DarkBlue, 5))
                {
                    graphics.DrawRectangle(pen, 50, 50, 700, 500);
                }
            }
            bitmap.Save(inputBmpPath);
        }

        // Insert the BMP image into a Word document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        Shape shape = builder.InsertImage(inputBmpPath);
        if (!shape.HasImage)
            throw new InvalidOperationException("Inserted shape does not contain an image.");

        // Save the document (optional, demonstrates lifecycle usage).
        const string docPath = "DocumentWithBmp.docx";
        doc.Save(docPath);

        // Extract each image, resize it to 640x480, and save the resized version.
        int imageIndex = 0;
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape s in shapeNodes)
        {
            if (!s.HasImage)
                continue;

            // Save the original image to a memory stream.
            using (var originalStream = new MemoryStream())
            {
                s.ImageData.Save(originalStream);
                originalStream.Position = 0; // Reset before reading.

                // Load the image into an Aspose.Drawing.Bitmap.
                using (var originalBmp = new Bitmap(originalStream))
                {
                    // Create a new bitmap with the target size.
                    using (var resizedBmp = new Bitmap(640, 480))
                    {
                        using (var graphics = Graphics.FromImage(resizedBmp))
                        {
                            graphics.Clear(Color.White);
                            // Draw the original image scaled to the new dimensions.
                            graphics.DrawImage(originalBmp, new Rectangle(0, 0, 640, 480));
                        }

                        // Save the resized BMP.
                        string resizedPath = $"resized_image_{imageIndex}.bmp";
                        resizedBmp.Save(resizedPath);
                        if (!File.Exists(resizedPath))
                            throw new InvalidOperationException($"Failed to save resized image '{resizedPath}'.");
                        imageIndex++;
                    }
                }
            }
        }

        // Ensure at least one image was processed.
        if (imageIndex == 0)
            throw new InvalidOperationException("No BMP images were extracted and resized.");
    }
}

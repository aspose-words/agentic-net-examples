using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a deterministic GIF image file.
        const string gifPath = "sample.gif";
        using (Bitmap bmp = new Bitmap(400, 400))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.LightBlue);
                g.DrawEllipse(new Pen(Color.DarkBlue, 5), 50, 50, 300, 300);
            }
            bmp.Save(gifPath, ImageFormat.Gif);
        }

        // Create a Word document and insert the GIF image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape gifShape = builder.InsertImage(gifPath);
        // Ensure the shape is appended to the paragraph (InsertImage already does this).
        doc.Save("DocumentWithGif.docx");

        // Extract GIF images, resize to 200x200, and convert to static PNG.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Gif)
                continue;

            // Save the original GIF image to a memory stream.
            using (MemoryStream gifStream = new MemoryStream())
            {
                shape.ImageData.Save(gifStream);
                gifStream.Position = 0; // Reset before reading.

                // Load the GIF into Aspose.Drawing.Bitmap.
                using (Bitmap originalBmp = new Bitmap(gifStream))
                {
                    // Create a new 200x200 bitmap for the resized PNG.
                    using (Bitmap resizedBmp = new Bitmap(200, 200))
                    {
                        using (Graphics graphics = Graphics.FromImage(resizedBmp))
                        {
                            graphics.Clear(Color.Transparent);
                            graphics.DrawImage(originalBmp, new Rectangle(0, 0, 200, 200));
                        }

                        string pngPath = $"extracted_{extractedCount}.png";
                        resizedBmp.Save(pngPath, ImageFormat.Png);

                        // Validate that the PNG file was created.
                        if (!File.Exists(pngPath))
                            throw new InvalidOperationException($"Failed to create PNG file: {pngPath}");

                        extractedCount++;
                    }
                }
            }
        }

        // Ensure at least one image was processed.
        if (extractedCount == 0)
            throw new InvalidOperationException("No GIF images were found and processed in the document.");
    }
}

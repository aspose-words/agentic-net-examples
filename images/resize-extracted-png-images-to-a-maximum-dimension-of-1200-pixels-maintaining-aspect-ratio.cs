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
        // Create a sample PNG image larger than 1200 pixels.
        const string originalImagePath = "sample.png";
        using (Bitmap originalBitmap = new Bitmap(2000, 1500))
        using (Graphics g = Graphics.FromImage(originalBitmap))
        {
            g.Clear(Color.White);
            originalBitmap.Save(originalImagePath, ImageFormat.Png);
        }

        // Create a Word document and insert the sample image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(originalImagePath);
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // Load the document and extract PNG images.
        Document loadedDoc = new Document(docPath);
        var pngShapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                 .Cast<Shape>()
                                 .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Png)
                                 .ToList();

        if (!pngShapes.Any())
            throw new InvalidOperationException("No PNG images were found in the document.");

        int imageIndex = 0;
        foreach (var shape in pngShapes)
        {
            // Save the image data to a memory stream.
            using (MemoryStream ms = new MemoryStream())
            {
                shape.ImageData.Save(ms);
                ms.Position = 0; // Reset stream before reading.

                // Load the image into Aspose.Drawing.Bitmap.
                using (Bitmap bitmap = new Bitmap(ms))
                {
                    int originalWidth = bitmap.Width;
                    int originalHeight = bitmap.Height;

                    // Determine scaling factor to keep max dimension <= 1200.
                    int maxDimension = Math.Max(originalWidth, originalHeight);
                    if (maxDimension <= 1200)
                    {
                        // No resizing needed; save the original image.
                        string unchangedPath = $"extracted_{imageIndex}.png";
                        bitmap.Save(unchangedPath, ImageFormat.Png);
                    }
                    else
                    {
                        double scale = 1200.0 / maxDimension;
                        int newWidth = (int)Math.Round(originalWidth * scale);
                        int newHeight = (int)Math.Round(originalHeight * scale);

                        using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                        using (Graphics graphics = Graphics.FromImage(resizedBitmap))
                        {
                            graphics.DrawImage(bitmap, 0, 0, newWidth, newHeight);
                            string resizedPath = $"resized_{imageIndex}.png";
                            resizedBitmap.Save(resizedPath, ImageFormat.Png);
                        }
                    }
                }
            }

            imageIndex++;
        }

        // Validate that at least one resized file was created.
        var outputFiles = Directory.GetFiles(Directory.GetCurrentDirectory(), "resized_*.png");
        if (!outputFiles.Any())
            throw new InvalidOperationException("No resized PNG images were produced.");
    }
}

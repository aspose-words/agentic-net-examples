using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string inputImagePath = "input.png";
        const string docPath = "DocumentWithImages.docx";

        // -------------------------------------------------
        // 1. Create a sample PNG image (800x400) locally
        // -------------------------------------------------
        Bitmap sampleBitmap = new Bitmap(800, 400);
        Graphics graphics = Graphics.FromImage(sampleBitmap);
        graphics.Clear(Color.White);
        // Draw a simple rectangle to make the image non‑empty
        graphics.DrawRectangle(new Pen(Color.Blue, 5), 50, 50, 700, 300);
        // Save the image as PNG
        sampleBitmap.Save(inputImagePath, ImageFormat.Png);
        graphics.Dispose();
        sampleBitmap.Dispose();

        // -------------------------------------------------
        // 2. Create a Word document and insert the PNG image
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        // Save the document containing the image
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract PNG images
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        var shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                 .OfType<Shape>()
                                 .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Png)
                                 .ToList();

        if (!shapeNodes.Any())
            throw new InvalidOperationException("No PNG images were found in the document.");

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            // Save the image data to a memory stream
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset for reading

                // Load the original bitmap from the stream
                using (Bitmap originalBitmap = new Bitmap(imageStream))
                {
                    // Desired fixed height
                    const int targetHeight = 600;
                    // Calculate proportional width
                    double scaleFactor = (double)targetHeight / originalBitmap.Height;
                    int targetWidth = (int)Math.Round(originalBitmap.Width * scaleFactor);

                    // Create a new bitmap with the target dimensions
                    using (Bitmap resizedBitmap = new Bitmap(targetWidth, targetHeight))
                    {
                        using (Graphics g = Graphics.FromImage(resizedBitmap))
                        {
                            // Draw the original image onto the resized bitmap
                            g.DrawImage(originalBitmap, 0, 0, targetWidth, targetHeight);
                        }

                        // Save the resized image to a deterministic file name
                        string resizedPath = $"resized_{imageIndex}.png";
                        resizedBitmap.Save(resizedPath, ImageFormat.Png);
                        imageIndex++;
                    }
                }
            }
        }

        // -------------------------------------------------
        // 4. Validation – ensure at least one resized file exists
        // -------------------------------------------------
        if (imageIndex == 0)
            throw new InvalidOperationException("No resized images were produced.");
    }
}

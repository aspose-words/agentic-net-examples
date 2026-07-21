using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string inputImagePath = "input.png";
        const string docPath = "DocumentWithImages.docx";

        // -------------------------------------------------
        // 1. Create a sample PNG image (800x400) using Aspose.Drawing
        // -------------------------------------------------
        Bitmap sampleBitmap = new Bitmap(800, 400);
        Graphics sampleGraphics = Graphics.FromImage(sampleBitmap);
        sampleGraphics.Clear(Color.White);
        // Draw a simple rectangle to make the image visible
        sampleGraphics.FillRectangle(
            new SolidBrush(Color.Blue),
            100, 100, 600, 200);
        sampleGraphics.Dispose();
        sampleBitmap.Save(inputImagePath, ImageFormat.Png);
        sampleBitmap.Dispose();

        // -------------------------------------------------
        // 2. Create a Word document and insert the sample image twice
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        builder.InsertParagraph(); // separate the images
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract PNG images, resizing each to 600px height
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        int resizedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Save the image data to a memory stream
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // reset before reading

                // Load the original bitmap
                using (Bitmap originalBitmap = new Bitmap(imageStream))
                {
                    // Calculate new dimensions preserving aspect ratio (height = 600px)
                    const int targetHeight = 600;
                    double scaleFactor = (double)targetHeight / originalBitmap.Height;
                    int targetWidth = (int)Math.Round(originalBitmap.Width * scaleFactor);

                    // Create a new bitmap with the target size
                    using (Bitmap resizedBitmap = new Bitmap(targetWidth, targetHeight))
                    {
                        using (Graphics graphics = Graphics.FromImage(resizedBitmap))
                        {
                            // High quality resizing
                            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            graphics.DrawImage(originalBitmap, 0, 0, targetWidth, targetHeight);
                        }

                        // Save the resized image to a deterministic file name
                        string resizedPath = $"resized_{imageIndex}.png";
                        resizedBitmap.Save(resizedPath, ImageFormat.Png);
                        if (!File.Exists(resizedPath))
                            throw new InvalidOperationException($"Failed to create resized image file: {resizedPath}");

                        resizedCount++;
                    }
                }
            }

            imageIndex++;
        }

        // Validation: ensure at least one image was resized
        if (resizedCount == 0)
            throw new InvalidOperationException("No PNG images were found and resized.");

        // Cleanup temporary files (optional)
        // File.Delete(inputImagePath);
        // File.Delete(docPath);
    }
}

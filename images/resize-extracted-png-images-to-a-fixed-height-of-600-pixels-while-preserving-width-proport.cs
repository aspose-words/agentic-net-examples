using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;   // For InterpolationMode

public class Program
{
    public static void Main()
    {
        // Define deterministic file names.
        const string inputImagePath = "input.png";
        const string docPath = "document.docx";

        // -------------------------------------------------
        // 1. Create a sample PNG image (800x400) using Aspose.Drawing.
        // -------------------------------------------------
        using (Bitmap bitmap = new Bitmap(800, 400))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Aspose.Drawing.Color.LightBlue);
            // Draw a simple rectangle for visual reference.
            using (Pen pen = new Pen(Aspose.Drawing.Color.DarkBlue, 5))
            {
                g.DrawRectangle(pen, 50, 50, 700, 300);
            }
            bitmap.Save(inputImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // 2. Create a Word document and insert the sample image.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        // Insert the same image a second time to demonstrate multiple extraction.
        builder.InsertParagraph();
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract PNG images.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Png)
                continue; // Process only PNG images.

            // Save the original image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0; // Reset before reading.

                // Load the image into Aspose.Drawing.Bitmap.
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    // Desired fixed height.
                    const int targetHeight = 600;
                    // Compute proportional width.
                    int targetWidth = (int)(originalBitmap.Width * (targetHeight / (double)originalBitmap.Height));

                    // Create a new bitmap with the target dimensions.
                    using (Bitmap resizedBitmap = new Bitmap(targetWidth, targetHeight))
                    using (Graphics graphics = Graphics.FromImage(resizedBitmap))
                    {
                        // High quality scaling.
                        graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        graphics.DrawImage(originalBitmap, 0, 0, targetWidth, targetHeight);

                        // Save the resized image to a deterministic file name.
                        string resizedPath = $"resized_{imageIndex}.png";
                        resizedBitmap.Save(resizedPath, ImageFormat.Png);
                        Console.WriteLine($"Resized image saved to: {resizedPath}");
                    }
                }
            }

            imageIndex++;
        }

        // -------------------------------------------------
        // 4. Validation: ensure at least one image was resized.
        // -------------------------------------------------
        if (imageIndex == 0)
            throw new InvalidOperationException("No PNG images were found and resized in the document.");
    }
}

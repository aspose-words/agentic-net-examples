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
        // Prepare deterministic file names.
        const string inputImagePath = "sample.png";
        const string docPath = "sample.docx";
        const string outputDocPath = "sample_sharpened.docx";

        // -------------------------------------------------
        // 1. Create a sample PNG image using Aspose.Drawing.
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill background with light gray.
                g.Clear(Color.LightGray);
                // Draw a simple red rectangle.
                g.FillRectangle(new SolidBrush(Color.Red), 50, 50, 100, 100);
            }
            // Save the image to a file so it can be inserted into the document.
            bitmap.Save(inputImagePath);
        }

        // -------------------------------------------------
        // 2. Create a Word document and insert the PNG image several times.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the image three times to have multiple shapes.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertImage(inputImagePath);
            builder.Writeln(); // Add a line break between images.
        }
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document, find all PNG images, sharpen them, and replace.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int processedCount = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Extract the image bytes.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the image into a Bitmap (Aspose.Drawing).
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    // Apply a simple sharpening kernel.
                    using (Bitmap sharpenedBitmap = ApplySharpenFilter(originalBitmap))
                    {
                        // Save the sharpened bitmap to a new stream.
                        using (MemoryStream sharpenedStream = new MemoryStream())
                        {
                            sharpenedBitmap.Save(sharpenedStream, ImageFormat.Png);
                            sharpenedStream.Position = 0;

                            // Replace the shape's image with the sharpened version.
                            shape.ImageData.SetImage(sharpenedStream);
                            processedCount++;
                        }
                    }
                }
            }
        }

        // Validate that at least one PNG image was processed.
        if (processedCount == 0)
            throw new InvalidOperationException("No PNG images were found to process.");

        // -------------------------------------------------
        // 4. Save the modified document.
        // -------------------------------------------------
        loadedDoc.Save(outputDocPath);
    }

    // -------------------------------------------------
    // Helper: Apply a 3x3 sharpening convolution kernel.
    // -------------------------------------------------
    private static Bitmap ApplySharpenFilter(Bitmap source)
    {
        int width = source.Width;
        int height = source.Height;
        Bitmap result = new Bitmap(width, height);

        // Sharpen kernel.
        int[,] kernel = {
            {  0, -1,  0 },
            { -1,  5, -1 },
            {  0, -1,  0 }
        };
        int kernelSize = 3;
        int offset = kernelSize / 2;

        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                int r = 0, g = 0, b = 0;

                for (int ky = -offset; ky <= offset; ky++)
                {
                    int py = y + ky;
                    if (py < 0 || py >= height) continue;

                    for (int kx = -offset; kx <= offset; kx++)
                    {
                        int px = x + kx;
                        if (px < 0 || px >= width) continue;

                        Color pixelColor = source.GetPixel(px, py);
                        int kernelValue = kernel[ky + offset, kx + offset];

                        r += pixelColor.R * kernelValue;
                        g += pixelColor.G * kernelValue;
                        b += pixelColor.B * kernelValue;
                    }
                }

                // Clamp color components to byte range.
                r = Math.Min(Math.Max(r, 0), 255);
                g = Math.Min(Math.Max(g, 0), 255);
                b = Math.Min(Math.Max(b, 0), 255);

                result.SetPixel(x, y, Color.FromArgb(r, g, b));
            }
        }

        return result;
    }
}

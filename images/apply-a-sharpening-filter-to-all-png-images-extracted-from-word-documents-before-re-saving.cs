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
        // Create a deterministic sample PNG image.
        const string sampleImagePath = "sample.png";
        const int imgWidth = 100;
        const int imgHeight = 100;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple black rectangle.
                g.FillRectangle(Brushes.Black, 20, 20, 60, 60);
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // Create a Word document and insert the sample PNG image twice.
        const string inputDocPath = "input.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        builder.Writeln(); // separate images
        builder.InsertImage(sampleImagePath);
        doc.Save(inputDocPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputDocPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int processedCount = 0;
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Extract the image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the image into a bitmap.
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    // Apply a sharpening filter.
                    Bitmap sharpenedBitmap = ApplySharpenFilter(originalBitmap);
                    // Save the sharpened bitmap to a new memory stream.
                    using (MemoryStream sharpenedStream = new MemoryStream())
                    {
                        sharpenedBitmap.Save(sharpenedStream, ImageFormat.Png);
                        sharpenedBitmap.Dispose();

                        sharpenedStream.Position = 0;
                        // Replace the shape's image with the sharpened version.
                        shape.ImageData.SetImage(sharpenedStream);
                    }

                    // Optionally, save the processed image to a file for verification.
                    string processedImagePath = $"processed-{imageIndex}.png";
                    using (FileStream fileOut = new FileStream(processedImagePath, FileMode.Create, FileAccess.Write))
                    {
                        sharpenedBitmap = new Bitmap(originalBitmap); // reload to save original size
                        sharpenedBitmap.Save(fileOut, ImageFormat.Png);
                        sharpenedBitmap.Dispose();
                    }
                }
            }

            processedCount++;
            imageIndex++;
        }

        // Validate that at least one PNG image was processed.
        if (processedCount == 0)
            throw new InvalidOperationException("No PNG images were found and processed in the document.");

        // Save the modified document.
        const string outputDocPath = "output.docx";
        loadedDoc.Save(outputDocPath);
    }

    // Applies a simple 3x3 sharpening kernel to the provided bitmap.
    private static Bitmap ApplySharpenFilter(Bitmap source)
    {
        int width = source.Width;
        int height = source.Height;
        Bitmap result = new Bitmap(width, height);

        // Copy original pixels to result (handles borders).
        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                result.SetPixel(x, y, source.GetPixel(x, y));
            }
        }

        // Sharpening kernel.
        int[,] kernel = {
            {  0, -1,  0 },
            { -1,  5, -1 },
            {  0, -1,  0 }
        };

        // Apply kernel to interior pixels.
        for (int y = 1; y < height - 1; y++)
        {
            for (int x = 1; x < width - 1; x++)
            {
                int r = 0, g = 0, b = 0;
                for (int ky = -1; ky <= 1; ky++)
                {
                    for (int kx = -1; kx <= 1; kx++)
                    {
                        Color pixel = source.GetPixel(x + kx, y + ky);
                        int factor = kernel[ky + 1, kx + 1];
                        r += pixel.R * factor;
                        g += pixel.G * factor;
                        b += pixel.B * factor;
                    }
                }
                // Clamp values to byte range.
                r = Math.Max(0, Math.Min(255, r));
                g = Math.Max(0, Math.Min(255, g));
                b = Math.Max(0, Math.Min(255, b));
                result.SetPixel(x, y, Color.FromArgb(r, g, b));
            }
        }

        return result;
    }
}

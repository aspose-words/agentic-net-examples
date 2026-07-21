using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare working directories.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);
        string inputImagePath = Path.Combine(workDir, "sample.png");
        string docPath = Path.Combine(workDir, "doc.docx");
        string archiveDir = Path.Combine(workDir, "SecureArchive");
        Directory.CreateDirectory(archiveDir);

        // Create a deterministic PNG image.
        CreateSamplePng(inputImagePath);

        // Create a document and insert the PNG image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // Extract images, convert each to grayscale BMP, and save to the secure folder.
        int imageIndex = 0;
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Save the shape's image to a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset before reading.

                // Load the image into Aspose.Drawing.Bitmap.
                using (Bitmap bitmap = new Bitmap(imageStream))
                {
                    // Convert the bitmap to grayscale.
                    ConvertToGrayscale(bitmap);

                    // Save the grayscale bitmap as BMP in the secure archive.
                    string bmpPath = Path.Combine(archiveDir, $"image_{imageIndex}.bmp");
                    bitmap.Save(bmpPath);
                }
            }

            imageIndex++;
        }

        // Validate that at least one image was processed.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");
    }

    // Creates a simple PNG image with a red rectangle on a white background.
    private static void CreateSamplePng(string filePath)
    {
        int width = 200;
        int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                using (Pen pen = new Pen(Color.Red, 5))
                {
                    graphics.DrawRectangle(pen, 20, 20, width - 40, height - 40);
                }
            }
            // Save the deterministic image.
            bitmap.Save(filePath);
        }
    }

    // Converts a bitmap to grayscale by adjusting each pixel.
    private static void ConvertToGrayscale(Bitmap bitmap)
    {
        for (int y = 0; y < bitmap.Height; y++)
        {
            for (int x = 0; x < bitmap.Width; x++)
            {
                Color pixel = bitmap.GetPixel(x, y);
                int gray = (int)(pixel.R * 0.3 + pixel.G * 0.59 + pixel.B * 0.11);
                Color grayColor = Color.FromArgb(pixel.A, gray, gray, gray);
                bitmap.SetPixel(x, y, grayColor);
            }
        }
    }
}

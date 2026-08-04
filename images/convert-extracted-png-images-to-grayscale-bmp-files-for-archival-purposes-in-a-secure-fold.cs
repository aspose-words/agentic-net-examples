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
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "Input");
        string archiveDir = Path.Combine(baseDir, "SecureArchive");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(archiveDir);

        // Create a sample PNG image
        string pngPath = Path.Combine(inputDir, "sample.png");
        CreateSamplePng(pngPath);

        // Create a Word document that contains the PNG image
        string docPath = Path.Combine(baseDir, "SampleDocument.docx");
        CreateDocumentWithImage(docPath, pngPath);

        // Load the document and extract PNG images
        Document doc = new Document(docPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;

            if (shape.ImageData.ImageType != ImageType.Png) continue;

            // Extract image bytes to a memory stream
            using (MemoryStream imgStream = new MemoryStream())
            {
                shape.ImageData.Save(imgStream);
                imgStream.Position = 0;

                // Load the PNG into Aspose.Drawing.Bitmap
                using (Bitmap bitmap = new Bitmap(imgStream))
                {
                    // Convert to grayscale
                    ConvertToGrayscale(bitmap);

                    // Save as BMP in the secure archive folder
                    string bmpFileName = Path.Combine(archiveDir, $"image_{imageIndex}.bmp");
                    bitmap.Save(bmpFileName, ImageFormat.Bmp);
                    if (!File.Exists(bmpFileName))
                        throw new InvalidOperationException($"Failed to create BMP file: {bmpFileName}");

                    imageIndex++;
                }
            }
        }

        // Validate that at least one BMP was created
        if (imageIndex == 0)
            throw new InvalidOperationException("No PNG images were found to convert.");

        // Example completed
        Console.WriteLine($"Converted {imageIndex} image(s) to grayscale BMP files in: {archiveDir}");
    }

    private static void CreateSamplePng(string filePath)
    {
        const int width = 200;
        const int height = 100;
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.White);
            // Draw a simple red rectangle
            using (Brush brush = new SolidBrush(Color.Red))
            {
                g.FillRectangle(brush, 20, 20, width - 40, height - 40);
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

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

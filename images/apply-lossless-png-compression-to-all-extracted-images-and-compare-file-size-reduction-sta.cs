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
        // Prepare folders
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample PNG image using Aspose.Drawing
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        CreateSamplePng(sampleImagePath, 200, 200);

        // 2. Create a Word document and insert the sample image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        string docPath = Path.Combine(artifactsDir, "document.docx");
        doc.Save(docPath);

        // 3. Extract images, recompress them losslessly, and gather statistics
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        var imageShapes = shapeNodes.OfType<Shape>().Where(s => s.HasImage).ToList();

        if (!imageShapes.Any())
            throw new InvalidOperationException("No images were found in the document.");

        long totalOriginalSize = 0;
        long totalCompressedSize = 0;
        int index = 0;

        foreach (Shape shape in imageShapes)
        {
            // Save original image
            string originalImagePath = Path.Combine(artifactsDir,
                $"extracted_{index}_original{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}");
            shape.ImageData.Save(originalImagePath);
            long originalSize = new FileInfo(originalImagePath).Length;
            totalOriginalSize += originalSize;

            // Load the original image into Aspose.Drawing.Bitmap
            using (FileStream originalStream = File.OpenRead(originalImagePath))
            {
                using (Bitmap bitmap = new Bitmap(originalStream))
                {
                    // Re‑save the bitmap as PNG (lossless compression)
                    string compressedImagePath = Path.Combine(artifactsDir,
                        $"extracted_{index}_compressed.png");
                    using (FileStream compressedStream = File.Create(compressedImagePath))
                    {
                        bitmap.Save(compressedStream, ImageFormat.Png);
                    }

                    long compressedSize = new FileInfo(compressedImagePath).Length;
                    totalCompressedSize += compressedSize;
                }
            }

            index++;
        }

        // 4. Output statistics
        Console.WriteLine($"Extracted {imageShapes.Count} image(s).");
        Console.WriteLine($"Total original size   : {totalOriginalSize} bytes");
        Console.WriteLine($"Total compressed size : {totalCompressedSize} bytes");

        if (totalOriginalSize > 0)
        {
            double reduction = 100.0 * (totalOriginalSize - totalCompressedSize) / totalOriginalSize;
            Console.WriteLine($"Size reduction        : {reduction:F2}%");
        }
        else
        {
            Console.WriteLine("Original size is zero, cannot compute reduction.");
        }
    }

    // Creates a deterministic PNG image using Aspose.Drawing
    private static void CreateSamplePng(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.White);
                // Draw a simple rectangle
                using (Pen pen = new Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }

            // Save the bitmap as PNG
            using (FileStream fs = File.Create(filePath))
            {
                bitmap.Save(fs, ImageFormat.Png);
            }
        }
    }
}

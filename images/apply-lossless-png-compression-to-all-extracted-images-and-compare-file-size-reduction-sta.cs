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
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a deterministic sample PNG image.
        string inputImagePath = Path.Combine(artifactsDir, "input.png");
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple red rectangle.
                using (Brush brush = new SolidBrush(Color.Red))
                {
                    g.FillRectangle(brush, 50, 50, 100, 100);
                }
            }
            bitmap.Save(inputImagePath, ImageFormat.Png);
        }

        // 2. Create a Word document and insert the sample image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        doc.Save(docPath);

        // 3. Load the document and extract all images.
        Document loadedDoc = new Document(docPath);
        var shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                  .Cast<Shape>()
                                  .Where(s => s.HasImage)
                                  .ToList();

        if (!shapeNodes.Any())
            throw new InvalidOperationException("No images were found in the document.");

        long totalOriginalSize = 0;
        long totalCompressedSize = 0;

        for (int i = 0; i < shapeNodes.Count; i++)
        {
            Shape shape = shapeNodes[i];

            // Save original image.
            string originalPath = Path.Combine(artifactsDir, $"original_{i}.png");
            using (FileStream originalFs = File.Create(originalPath))
            {
                shape.ImageData.Save(originalFs);
            }
            long originalSize = new FileInfo(originalPath).Length;
            totalOriginalSize += originalSize;

            // Load the original image into Aspose.Drawing.Bitmap.
            using (FileStream originalFsRead = File.OpenRead(originalPath))
            using (Bitmap bmp = new Bitmap(originalFsRead))
            {
                // Re‑save the bitmap as PNG (lossless compression).
                string compressedPath = Path.Combine(artifactsDir, $"compressed_{i}.png");
                using (FileStream compressedFs = File.Create(compressedPath))
                {
                    bmp.Save(compressedFs, ImageFormat.Png);
                }
                long compressedSize = new FileInfo(compressedPath).Length;
                totalCompressedSize += compressedSize;

                // Output size comparison for this image.
                Console.WriteLine($"Image {i}: Original = {originalSize} bytes, Compressed = {compressedSize} bytes, Reduction = {originalSize - compressedSize} bytes");
            }
        }

        // 4. Output overall statistics.
        Console.WriteLine($"Total original size: {totalOriginalSize} bytes");
        Console.WriteLine($"Total compressed size: {totalCompressedSize} bytes");
        Console.WriteLine($"Total reduction: {totalOriginalSize - totalCompressedSize} bytes");
    }
}

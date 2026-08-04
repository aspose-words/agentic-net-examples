using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Directories for artifacts
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a deterministic sample PNG image.
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(200, 200))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.White);
            // Draw a simple rectangle.
            g.FillRectangle(new SolidBrush(Color.FromArgb(255, 0, 120, 215)), 20, 20, 160, 160);
            bitmap.Save(sampleImagePath);
        }

        // 2. Create a Word document and insert the sample image multiple times.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Images before compression:");
        builder.InsertImage(sampleImagePath);
        builder.InsertParagraph();
        builder.InsertImage(sampleImagePath);
        string originalDocPath = Path.Combine(artifactsDir, "Original.docx");
        doc.Save(originalDocPath);

        // 3. Load the document (demonstrating load usage).
        Document loadedDoc = new Document(originalDocPath);

        // 4. Extract all images from the document.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        bool anyImageExtracted = false;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            anyImageExtracted = true;

            // Original extracted image.
            string originalImagePath = Path.Combine(artifactsDir, $"extracted_original_{imageIndex}.png");
            shape.ImageData.Save(originalImagePath);

            // 5. Apply lossless PNG compression.
            // Load the original image using Aspose.Drawing, then re‑save it as PNG.
            string compressedImagePath = Path.Combine(artifactsDir, $"compressed_{imageIndex}.png");
            using (FileStream originalStream = File.OpenRead(originalImagePath))
            using (MemoryStream ms = new MemoryStream())
            {
                // Ensure the stream is at the beginning.
                originalStream.Position = 0;
                originalStream.CopyTo(ms);
                ms.Position = 0;

                using (Image img = Image.FromStream(ms))
                {
                    // Re‑save with PNG format (lossless compression).
                    img.Save(compressedImagePath, ImageFormat.Png);
                }
            }

            // 6. Compare file sizes.
            long originalSize = new FileInfo(originalImagePath).Length;
            long compressedSize = new FileInfo(compressedImagePath).Length;
            double reduction = originalSize == 0 ? 0 : 100.0 * (originalSize - compressedSize) / originalSize;

            Console.WriteLine($"Image {imageIndex}:");
            Console.WriteLine($"  Original size  : {originalSize} bytes");
            Console.WriteLine($"  Compressed size: {compressedSize} bytes");
            Console.WriteLine($"  Size reduction : {reduction:F2}%");
            Console.WriteLine();

            imageIndex++;
        }

        // Validation: ensure at least one image was processed.
        if (!anyImageExtracted)
            throw new InvalidOperationException("No images were extracted from the document.");

        // 7. Optionally, save the document after processing (not required for this task).
        // loadedDoc.Save(Path.Combine(artifactsDir, "Processed.docx"));
    }
}

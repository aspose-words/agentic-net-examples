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

        // -----------------------------------------------------------------
        // 1. Create a sample PNG image using Aspose.Drawing.
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(artifactsDir, "input.png");
        using (Bitmap bitmap = new Bitmap(200, 200))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            using (Pen pen = new Pen(Color.Black, 2))
            {
                graphics.DrawRectangle(pen, 10, 10, 180, 180);
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 2. Build a Word document that contains the sample image twice.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        builder.InsertParagraph();
        builder.InsertImage(sampleImagePath);
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Extract all images from the document.
        // -----------------------------------------------------------------
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        var imageShapes = shapeNodes.OfType<Shape>().Where(s => s.HasImage).ToList();

        if (imageShapes.Count == 0)
            throw new InvalidOperationException("No images were found in the document.");

        int imageIndex = 0;
        foreach (Shape shape in imageShapes)
        {
            // Save the original extracted image.
            string originalPath = Path.Combine(artifactsDir, $"original_{imageIndex}.png");
            shape.ImageData.Save(originalPath);

            // -----------------------------------------------------------------
            // 4. Re‑encode the image as lossless PNG (Aspose.Drawing does this by default).
            // -----------------------------------------------------------------
            string compressedPath = Path.Combine(artifactsDir, $"compressed_{imageIndex}.png");
            using (Image img = Image.FromFile(originalPath))
            {
                img.Save(compressedPath, ImageFormat.Png);
            }

            // -----------------------------------------------------------------
            // 5. Compare file sizes and output statistics.
            // -----------------------------------------------------------------
            long originalSize = new FileInfo(originalPath).Length;
            long compressedSize = new FileInfo(compressedPath).Length;
            double reductionPercent = originalSize == 0
                ? 0
                : (originalSize - compressedSize) * 100.0 / originalSize;

            Console.WriteLine($"Image {imageIndex}: Original = {originalSize} bytes, " +
                              $"Compressed = {compressedSize} bytes, " +
                              $"Reduction = {reductionPercent:0.##}%");

            imageIndex++;
        }
    }
}

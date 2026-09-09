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
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        const int imgWidth = 200;
        const int imgHeight = 200;

        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            // Fill background.
            graphics.Clear(Aspose.Drawing.Color.White);

            // Draw a simple red ellipse.
            using (Pen pen = new Pen(Aspose.Drawing.Color.Red, 5))
            {
                graphics.DrawEllipse(pen, 20, 20, imgWidth - 40, imgHeight - 40);
            }

            // Save the bitmap as PNG.
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 2. Create a Word document and insert the sample image several times.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image twice with a paragraph between them.
        builder.InsertImage(sampleImagePath);
        builder.InsertParagraph();
        builder.InsertImage(sampleImagePath);

        // Save the original document.
        string originalDocPath = Path.Combine(artifactsDir, "original.docx");
        doc.Save(originalDocPath);

        // -----------------------------------------------------------------
        // 3. Extract all images, recompress them losslessly as PNG, and
        //    compare file size statistics.
        // -----------------------------------------------------------------
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        bool anyImageFound = false;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            anyImageFound = true;

            // Save original image bytes to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                long originalSize = originalStream.Length;

                // Rewind the stream before loading the image.
                originalStream.Position = 0;

                // Load the image with Aspose.Drawing.
                using (Aspose.Drawing.Image img = Aspose.Drawing.Image.FromStream(originalStream))
                {
                    // Re‑encode the image as PNG (lossless).
                    using (MemoryStream compressedStream = new MemoryStream())
                    {
                        img.Save(compressedStream, ImageFormat.Png);
                        long compressedSize = compressedStream.Length;

                        // Write the compressed image to a file for verification.
                        string compressedPath = Path.Combine(artifactsDir, $"compressed_{imageIndex}.png");
                        File.WriteAllBytes(compressedPath, compressedStream.ToArray());

                        // Output statistics.
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
        }

        // Validate that at least one image was processed.
        if (!anyImageFound)
            throw new InvalidOperationException("No images were found in the document to process.");
    }
}

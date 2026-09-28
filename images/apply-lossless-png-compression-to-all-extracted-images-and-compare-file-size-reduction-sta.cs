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
        // Step 1: Create a deterministic sample PNG image.
        const string sampleImagePath = "sample.png";
        CreateSamplePng(sampleImagePath, 200, 200);

        // Step 2: Create a Word document and insert the sample image.
        const string docPath = "DocumentWithImages.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // Step 3: Load the document (optional, we already have it) and extract images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        bool anyImageExtracted = false;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage) continue;

            anyImageExtracted = true;
            ImageData imgData = shape.ImageData;

            // Save original extracted image to a temporary file.
            string originalImagePath = $"extracted-{imageIndex}.png";
            using (MemoryStream originalStream = new MemoryStream())
            {
                imgData.Save(originalStream);
                originalStream.Position = 0;
                File.WriteAllBytes(originalImagePath, originalStream.ToArray());
            }

            // If the image is PNG, apply lossless compression by re‑saving it.
            if (imgData.ImageType == ImageType.Png)
            {
                string compressedImagePath = $"compressed-{imageIndex}.png";

                // Load the original PNG into Aspose.Drawing.Bitmap.
                using (MemoryStream ms = new MemoryStream(File.ReadAllBytes(originalImagePath)))
                {
                    using (Bitmap bitmap = new Bitmap(ms))
                    {
                        // Re‑save the bitmap as PNG (lossless). This may reduce file size.
                        bitmap.Save(compressedImagePath, ImageFormat.Png);
                    }
                }

                // Compare file sizes.
                long originalSize = new FileInfo(originalImagePath).Length;
                long compressedSize = new FileInfo(compressedImagePath).Length;
                Console.WriteLine($"Image {imageIndex}: Original PNG size = {originalSize} bytes, " +
                                  $"Compressed PNG size = {compressedSize} bytes, " +
                                  $"Reduction = {originalSize - compressedSize} bytes.");
            }
            else
            {
                Console.WriteLine($"Image {imageIndex}: Not a PNG image, skipped compression.");
            }

            imageIndex++;
        }

        if (!anyImageExtracted)
            throw new InvalidOperationException("No images were extracted from the document.");

        Console.WriteLine("Processing completed.");
    }

    private static void CreateSamplePng(string filePath, int width, int height)
    {
        // Create a bitmap and draw deterministic content.
        Bitmap bitmap = new Bitmap(width, height);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        // Draw a simple red rectangle.
        using (Pen pen = new Pen(Color.Red, 5))
        {
            graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
        }
        // Save the bitmap as PNG.
        bitmap.Save(filePath, ImageFormat.Png);
        // Clean up.
        graphics.Dispose();
        bitmap.Dispose();
    }
}

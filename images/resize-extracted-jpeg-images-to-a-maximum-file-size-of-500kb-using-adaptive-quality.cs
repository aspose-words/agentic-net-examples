using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    // Maximum allowed size for a JPEG image (500 KB).
    private const long MaxJpegSizeBytes = 500 * 1024;

    public static void Main()
    {
        // Ensure the output directories exist.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);
        string inputImagePath = "input.jpg";

        // ------------------------------------------------------------
        // 1. Create a sample high‑resolution JPEG image using Aspose.Drawing.
        // ------------------------------------------------------------
        const int imgWidth = 2000;
        const int imgHeight = 1500;
        using (var bitmap = new Bitmap(imgWidth, imgHeight))
        using (var graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            using (var pen = new Pen(Color.Blue, 10))
            {
                graphics.DrawRectangle(pen, 100, 100, imgWidth - 200, imgHeight - 200);
            }
            // Save with the highest quality (100) to guarantee a large file.
            bitmap.Save(inputImagePath, ImageFormat.Jpeg);
        }

        // ------------------------------------------------------------
        // 2. Create a Word document and insert the sample image.
        // ------------------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        string docPath = Path.Combine(artifactsDir, "Original.docx");
        doc.Save(docPath);

        // ------------------------------------------------------------
        // 3. Extract JPEG images, recompress them adaptively, and replace.
        // ------------------------------------------------------------
        var shapes = doc.GetChildNodes(NodeType.Shape, true)
                        .Cast<Shape>()
                        .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Jpeg)
                        .ToList();

        if (!shapes.Any())
            throw new InvalidOperationException("No JPEG images were found in the document.");

        foreach (var shape in shapes)
        {
            // Get the original image bytes.
            byte[] originalBytes = shape.ImageData.ToByteArray();

            // Adaptive quality loop.
            int quality = 100;
            byte[] compressedBytes = null;

            while (quality >= 10)
            {
                // Create a temporary document containing only this image.
                var tempDoc = new Document();
                var tempBuilder = new DocumentBuilder(tempDoc);
                tempBuilder.InsertImage(originalBytes);

                // Save the temporary document as a JPEG image with the current quality.
                var imgOptions = new ImageSaveOptions(SaveFormat.Jpeg)
                {
                    JpegQuality = quality
                };
                using (var outStream = new MemoryStream())
                {
                    tempDoc.Save(outStream, imgOptions);
                    compressedBytes = outStream.ToArray();
                }

                // Check the size.
                if (compressedBytes.Length <= MaxJpegSizeBytes)
                    break; // Desired size achieved.

                // Reduce quality and try again.
                quality -= 10;
            }

            // If even the lowest quality exceeds the limit, keep the smallest we obtained.
            if (compressedBytes == null)
                throw new InvalidOperationException("Failed to compress image.");

            // Replace the image in the original shape with the compressed version.
            using (var ms = new MemoryStream(compressedBytes))
            {
                ms.Position = 0;
                shape.ImageData.SetImage(ms);
            }
        }

        // ------------------------------------------------------------
        // 4. Save the modified document.
        // ------------------------------------------------------------
        string resultDocPath = Path.Combine(artifactsDir, "Resized.docx");
        doc.Save(resultDocPath);

        // ------------------------------------------------------------
        // 5. Validation – ensure the output document exists and images are within size limit.
        // ------------------------------------------------------------
        if (!File.Exists(resultDocPath))
            throw new FileNotFoundException("Result document was not created.", resultDocPath);

        var resultDoc = new Document(resultDocPath);
        var resultShapes = resultDoc.GetChildNodes(NodeType.Shape, true)
                                   .Cast<Shape>()
                                   .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Jpeg);

        foreach (var shape in resultShapes)
        {
            byte[] bytes = shape.ImageData.ToByteArray();
            if (bytes.Length > MaxJpegSizeBytes)
                throw new InvalidOperationException($"An image exceeds the size limit: {bytes.Length} bytes.");
        }

        // Successful completion – no console output required.
    }
}

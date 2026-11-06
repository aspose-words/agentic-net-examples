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
        // Create a sample JPEG image.
        const string sampleImagePath = "sample.jpg";
        CreateSampleJpeg(sampleImagePath);

        // Create a Word document and insert the sample image.
        const string originalDocPath = "original.docx";
        CreateDocumentWithImage(originalDocPath, sampleImagePath);

        // Load the document, blur all JPEG images, and save a new document.
        const string blurredDocPath = "blurred.docx";
        ApplyBlurToJpegImages(originalDocPath, blurredDocPath);

        // Validate output.
        if (!File.Exists(blurredDocPath))
            throw new Exception("Blurred document was not created.");
    }

    private static void CreateSampleJpeg(string path)
    {
        int width = 200;
        int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a simple red rectangle.
                using (SolidBrush brush = new SolidBrush(Aspose.Drawing.Color.Red))
                {
                    g.FillRectangle(brush, 50, 50, 100, 100);
                }
            }
            // Save as JPEG.
            bitmap.Save(path, ImageFormat.Jpeg);
        }

        if (!File.Exists(path))
            throw new Exception("Sample JPEG image was not created.");
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
        if (!File.Exists(docPath))
            throw new Exception("Document with image was not created.");
    }

    private static void ApplyBlurToJpegImages(string inputDocPath, string outputDocPath)
    {
        Document doc = new Document(inputDocPath);
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int processedCount = 0;

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Extract image to memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load bitmap.
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    // Apply blur.
                    using (Bitmap blurredBitmap = ApplyAverageBlur(originalBitmap))
                    {
                        // Save blurred bitmap to stream.
                        using (MemoryStream blurredStream = new MemoryStream())
                        {
                            blurredBitmap.Save(blurredStream, ImageFormat.Jpeg);
                            blurredStream.Position = 0;

                            // Replace image data in the shape.
                            shape.ImageData.SetImage(blurredStream);
                        }
                    }
                }
            }

            processedCount++;
        }

        if (processedCount == 0)
            throw new Exception("No JPEG images were found to process.");

        doc.Save(outputDocPath);
    }

    private static Bitmap ApplyAverageBlur(Bitmap source)
    {
        int width = source.Width;
        int height = source.Height;
        Bitmap result = new Bitmap(width, height);

        // Simple 3x3 average blur kernel.
        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                int rSum = 0, gSum = 0, bSum = 0, count = 0;

                for (int ky = -1; ky <= 1; ky++)
                {
                    int ny = y + ky;
                    if (ny < 0 || ny >= height) continue;

                    for (int kx = -1; kx <= 1; kx++)
                    {
                        int nx = x + kx;
                        if (nx < 0 || nx >= width) continue;

                        Aspose.Drawing.Color pixel = source.GetPixel(nx, ny);
                        rSum += pixel.R;
                        gSum += pixel.G;
                        bSum += pixel.B;
                        count++;
                    }
                }

                int r = rSum / count;
                int g = gSum / count;
                int b = bSum / count;
                result.SetPixel(x, y, Aspose.Drawing.Color.FromArgb(r, g, b));
            }
        }

        return result;
    }
}

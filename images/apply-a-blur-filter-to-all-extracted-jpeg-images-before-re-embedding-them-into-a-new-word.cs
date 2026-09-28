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
        string sampleImagePath = "sample.jpg";
        CreateSampleJpeg(sampleImagePath);

        // Create a Word document that contains the sample image.
        string originalDocPath = "original.docx";
        CreateDocumentWithImage(originalDocPath, sampleImagePath);

        // Process the document: extract each JPEG, apply blur, and re‑embed.
        string blurredDocPath = "blurred.docx";
        ProcessDocumentImages(originalDocPath, blurredDocPath);

        // Validate that the output document was created.
        if (!File.Exists(blurredDocPath))
            throw new Exception("Blurred document was not created.");
    }

    // Generates a deterministic JPEG image using Aspose.Drawing.
    private static void CreateSampleJpeg(string path)
    {
        int width = 200;
        int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Red, 5))
                {
                    g.DrawRectangle(pen, 20, 20, width - 40, height - 40);
                }
            }
            bitmap.Save(path, ImageFormat.Jpeg);
        }
    }

    // Inserts the given image into a new Word document.
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Extracts JPEG images, applies a blur filter, and replaces them in a new document.
    private static void ProcessDocumentImages(string inputDocPath, string outputDocPath)
    {
        Document doc = new Document(inputDocPath);
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Extract the image to a memory stream.
            using (MemoryStream extractStream = new MemoryStream())
            {
                shape.ImageData.Save(extractStream);
                extractStream.Position = 0;

                // Load the image into a bitmap.
                using (Bitmap originalBitmap = new Bitmap(extractStream))
                {
                    // Apply a simple box blur.
                    using (Bitmap blurredBitmap = ApplyBoxBlur(originalBitmap, 1))
                    {
                        // Save the blurred bitmap to a new stream.
                        using (MemoryStream blurredStream = new MemoryStream())
                        {
                            blurredBitmap.Save(blurredStream, ImageFormat.Jpeg);
                            blurredStream.Position = 0;

                            // Replace the shape's image data with the blurred image.
                            shape.ImageData.SetImage(blurredStream);
                        }
                    }
                }
            }
        }
        doc.Save(outputDocPath);
    }

    // Simple box blur implementation.
    private static Bitmap ApplyBoxBlur(Bitmap source, int radius)
    {
        int width = source.Width;
        int height = source.Height;
        Bitmap result = new Bitmap(width, height);
        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                int a = 0, r = 0, g = 0, b = 0;
                int count = 0;
                for (int ky = -radius; ky <= radius; ky++)
                {
                    int ny = y + ky;
                    if (ny < 0 || ny >= height) continue;
                    for (int kx = -radius; kx <= radius; kx++)
                    {
                        int nx = x + kx;
                        if (nx < 0 || nx >= width) continue;
                        Aspose.Drawing.Color pixel = source.GetPixel(nx, ny);
                        a += pixel.A;
                        r += pixel.R;
                        g += pixel.G;
                        b += pixel.B;
                        count++;
                    }
                }
                result.SetPixel(x, y, Aspose.Drawing.Color.FromArgb(a / count, r / count, g / count, b / count));
            }
        }
        return result;
    }
}

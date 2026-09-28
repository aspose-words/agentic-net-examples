using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample JPEG image.
        string sampleImagePath = "sample.jpg";
        CreateSampleJpeg(sampleImagePath);

        // Build a document and insert the JPEG image twice.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        builder.Writeln();
        builder.InsertImage(sampleImagePath);
        string originalDocPath = "original.docx";
        doc.Save(originalDocPath);

        // Load the document for processing.
        Document loadDoc = new Document(originalDocPath);
        NodeCollection shapeNodes = loadDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Jpeg)
            {
                // Extract the JPEG image to a memory stream.
                using (MemoryStream ms = new MemoryStream())
                {
                    shape.ImageData.Save(ms);
                    ms.Position = 0;

                    // Load the image into a bitmap.
                    using (Bitmap originalBitmap = new Bitmap(ms))
                    {
                        // Apply a simple horizontal motion blur.
                        using (Bitmap blurredBitmap = ApplyMotionBlur(originalBitmap, 10))
                        {
                            // Save the blurred image to a temporary file.
                            string blurredPath = $"blurred_{imageIndex}.jpg";
                            blurredBitmap.Save(blurredPath);
                            if (!File.Exists(blurredPath))
                                throw new Exception($"Blurred image not saved: {blurredPath}");

                            // Replace the shape's image with the blurred version.
                            shape.ImageData.SetImage(blurredPath);
                        }
                    }
                }
                imageIndex++;
            }
        }

        // Save the final document.
        string outputDocPath = "output.docx";
        loadDoc.Save(outputDocPath);
        if (!File.Exists(outputDocPath))
            throw new Exception("Output document was not created.");
    }

    // Creates a simple deterministic JPEG image.
    private static void CreateSampleJpeg(string path)
    {
        int width = 200;
        int height = 100;
        Bitmap bitmap = new Bitmap(width, height);
        Graphics g = Graphics.FromImage(bitmap);
        g.Clear(Color.White);
        using (SolidBrush brush = new SolidBrush(Color.Blue))
        {
            g.FillRectangle(brush, 20, 20, 160, 60);
        }
        g.Dispose();
        bitmap.Save(path);
        bitmap.Dispose();

        if (!File.Exists(path))
            throw new Exception($"Sample JPEG not created: {path}");
    }

    // Applies a basic horizontal motion blur to a bitmap.
    private static Bitmap ApplyMotionBlur(Bitmap src, int length)
    {
        int width = src.Width;
        int height = src.Height;
        Bitmap dest = new Bitmap(width, height);

        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                int rSum = 0, gSum = 0, bSum = 0, count = 0;
                for (int k = 0; k < length; k++)
                {
                    int nx = x + k;
                    if (nx >= width) break;
                    Color col = src.GetPixel(nx, y);
                    rSum += col.R;
                    gSum += col.G;
                    bSum += col.B;
                    count++;
                }
                int r = rSum / count;
                int g = gSum / count;
                int b = bSum / count;
                Color avg = Color.FromArgb(r, g, b);
                dest.SetPixel(x, y, avg);
            }
        }

        return dest;
    }
}

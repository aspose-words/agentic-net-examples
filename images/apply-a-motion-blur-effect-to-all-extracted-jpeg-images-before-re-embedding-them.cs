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
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create two sample JPEG images.
        string img1Path = Path.Combine(artifactsDir, "sample1.jpg");
        string img2Path = Path.Combine(artifactsDir, "sample2.jpg");
        CreateSampleJpeg(img1Path, Color.Red);
        CreateSampleJpeg(img2Path, Color.Blue);

        // Build a document that contains the sample images.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(img1Path);
        builder.InsertParagraph();
        builder.InsertImage(img2Path);
        string originalDocPath = Path.Combine(artifactsDir, "original.docx");
        doc.Save(originalDocPath);

        // Load the document and apply a motion‑blur effect to every JPEG image.
        Document loadedDoc = new Document(originalDocPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Jpeg)
                ApplyMotionBlur(shape);
        }

        // Save the modified document.
        string blurredDocPath = Path.Combine(artifactsDir, "blurred.docx");
        loadedDoc.Save(blurredDocPath);

        // Simple validation.
        if (!File.Exists(blurredDocPath))
            throw new Exception("The blurred document was not saved.");
    }

    // Creates a deterministic JPEG image with a solid colored ellipse.
    private static void CreateSampleJpeg(string filePath, Color fillColor)
    {
        const int width = 200;
        const int height = 200;

        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.White);
            using (SolidBrush brush = new SolidBrush(fillColor))
            {
                g.FillEllipse(brush, 20, 20, width - 40, height - 40);
            }
            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }

    // Applies a simple horizontal motion‑blur by drawing several shifted copies of the image.
    private static void ApplyMotionBlur(Shape shape)
    {
        // Retrieve the original image bytes.
        byte[] originalBytes = shape.ImageData.ToByteArray();

        using (MemoryStream srcStream = new MemoryStream(originalBytes))
        using (Bitmap srcBitmap = new Bitmap(srcStream))
        {
            int w = srcBitmap.Width;
            int h = srcBitmap.Height;

            using (Bitmap dstBitmap = new Bitmap(w, h))
            using (Graphics g = Graphics.FromImage(dstBitmap))
            {
                // Transparent background.
                g.Clear(Color.Transparent);

                // Number of shifted copies – larger value = stronger blur.
                const int blurLength = 10;

                // Draw the source image multiple times, each shifted one pixel to the right.
                for (int i = 0; i < blurLength; i++)
                {
                    g.DrawImage(srcBitmap, i, 0, w, h);
                }

                // Save the blurred image to a memory stream.
                using (MemoryStream outStream = new MemoryStream())
                {
                    dstBitmap.Save(outStream, ImageFormat.Jpeg);
                    outStream.Position = 0;

                    // Replace the shape's image with the blurred version.
                    shape.ImageData.SetImage(outStream);
                }
            }
        }
    }
}

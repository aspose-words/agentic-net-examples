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
        CreateSamplePng(sampleImagePath);

        // Step 2: Create a Word document and insert the sample image twice.
        const string docPath = "sample.docx";
        CreateDocumentWithImages(docPath, sampleImagePath);

        // Step 3: Load the document and extract PNG images, applying contrast enhancement.
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;
        int index = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset before reading.

                // Load the image using Aspose.Drawing.Bitmap.
                using (Bitmap original = new Bitmap(imageStream))
                {
                    // Convert to a non‑indexed pixel format (32bpp ARGB) to allow SetPixel.
                    using (Bitmap bitmap = new Bitmap(original.Width, original.Height, PixelFormat.Format32bppArgb))
                    {
                        using (Graphics g = Graphics.FromImage(bitmap))
                        {
                            g.DrawImage(original, 0, 0, original.Width, original.Height);
                        }

                        // Apply contrast enhancement (contrast value: 50).
                        AdjustContrast(bitmap, 50f);

                        string outputPath = $"extracted_contrast_{index}.png";
                        bitmap.Save(outputPath, ImageFormat.Png);
                        extractedCount++;
                        index++;
                    }
                }
            }
        }

        // Validation: ensure at least one image was processed.
        if (extractedCount == 0)
            throw new InvalidOperationException("No PNG images were extracted and processed.");
    }

    // Creates a simple deterministic PNG image.
    private static void CreateSamplePng(string path)
    {
        const int width = 200;
        const int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple blue ellipse.
                using (SolidBrush brush = new SolidBrush(Color.Blue))
                {
                    g.FillEllipse(brush, 20, 20, width - 40, height - 40);
                }
            }
            bitmap.Save(path, ImageFormat.Png);
        }
    }

    // Creates a Word document and inserts the same image twice.
    private static void CreateDocumentWithImages(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        builder.Writeln(); // Add a line break between images.
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Adjusts contrast of the bitmap. Contrast range: -255 to 255.
    private static void AdjustContrast(Bitmap bitmap, float contrast)
    {
        // Clamp contrast to valid range.
        if (contrast < -255f) contrast = -255f;
        if (contrast > 255f) contrast = 255f;

        double factor = (259 * (contrast + 255)) / (255.0 * (259 - contrast));

        for (int y = 0; y < bitmap.Height; y++)
        {
            for (int x = 0; x < bitmap.Width; x++)
            {
                Color original = bitmap.GetPixel(x, y);
                int r = Truncate((int)(factor * (original.R - 128) + 128));
                int g = Truncate((int)(factor * (original.G - 128) + 128));
                int b = Truncate((int)(factor * (original.B - 128) + 128));
                Color adjusted = Color.FromArgb(original.A, r, g, b);
                bitmap.SetPixel(x, y, adjusted);
            }
        }
    }

    private static int Truncate(int value)
    {
        if (value < 0) return 0;
        if (value > 255) return 255;
        return value;
    }
}

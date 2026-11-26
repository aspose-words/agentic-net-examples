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

        // Create a Word document and insert the JPEG image.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // Reload the document for extraction.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        int index = 0;
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Get image bytes.
            byte[] imageBytes = shape.ImageData.ImageBytes;
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                using (Bitmap bitmap = new Bitmap(ms))
                {
                    // Increase saturation by 20% (simple RGB scaling).
                    for (int y = 0; y < bitmap.Height; y++)
                    {
                        for (int x = 0; x < bitmap.Width; x++)
                        {
                            Color original = bitmap.GetPixel(x, y);
                            int r = Math.Min((int)(original.R * 1.2), 255);
                            int g = Math.Min((int)(original.G * 1.2), 255);
                            int b = Math.Min((int)(original.B * 1.2), 255);
                            Color adjusted = Color.FromArgb(original.A, r, g, b);
                            bitmap.SetPixel(x, y, adjusted);
                        }
                    }

                    // Save the adjusted image.
                    string outputPath = $"extracted_{index}_saturated.jpg";
                    bitmap.Save(outputPath, ImageFormat.Jpeg);
                    extractedCount++;
                    index++;
                }
            }
        }

        if (extractedCount == 0)
            throw new InvalidOperationException("No JPEG images were extracted and processed.");

        // Optional clean‑up.
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }

    private static void CreateSampleJpeg(string path)
    {
        const int width = 200;
        const int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 100, 150, 200)))
                {
                    g.FillEllipse(brush, 20, 20, width - 40, height - 40);
                }
            }
            bitmap.Save(path, ImageFormat.Jpeg);
        }
    }
}

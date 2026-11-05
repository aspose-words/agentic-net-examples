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

        // Create a DOCX file that contains the JPEG image.
        string docPath = "input.docx";
        CreateDocumentWithImages(docPath, sampleImagePath);

        // Extract JPEG images, apply grayscale filter, and save them.
        ExtractAndGrayscaleImages(docPath);
    }

    private static void CreateSampleJpeg(string path)
    {
        int width = 200;
        int height = 100;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                using (SolidBrush brush = new SolidBrush(Color.Blue))
                {
                    g.FillRectangle(brush, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(path, ImageFormat.Jpeg);
        }
    }

    private static void CreateDocumentWithImages(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        builder.Writeln();
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    private static void ExtractAndGrayscaleImages(string docPath)
    {
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int savedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            byte[] imageBytes = shape.ImageData.ImageBytes;
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0;
                using (Bitmap bitmap = new Bitmap(ms))
                {
                    ApplyGrayscale(bitmap);
                    string outputPath = $"extracted-{savedCount}.jpg";
                    bitmap.Save(outputPath, ImageFormat.Jpeg);
                    savedCount++;
                }
            }
        }

        if (savedCount == 0)
            throw new InvalidOperationException("No JPEG images were extracted and saved.");
    }

    private static void ApplyGrayscale(Bitmap bitmap)
    {
        for (int y = 0; y < bitmap.Height; y++)
        {
            for (int x = 0; x < bitmap.Width; x++)
            {
                Color original = bitmap.GetPixel(x, y);
                int gray = (int)(original.R * 0.3 + original.G * 0.59 + original.B * 0.11);
                Color grayColor = Color.FromArgb(original.A, gray, gray, gray);
                bitmap.SetPixel(x, y, grayColor);
            }
        }
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare output directories.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(artifactsDir);
        string imagesDir = Path.Combine(artifactsDir, "Images");
        Directory.CreateDirectory(imagesDir);

        // Create sample PNG and JPEG images.
        string pngPath = Path.Combine(imagesDir, "sample.png");
        CreateSamplePng(pngPath);
        string jpegPath = Path.Combine(imagesDir, "sample.jpg");
        CreateSampleJpeg(jpegPath);

        // Build a new document and insert the images.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(pngPath);
        builder.InsertParagraph(); // optional spacing between images
        builder.InsertImage(jpegPath);

        // Save the document as a single PDF file.
        string pdfPath = Path.Combine(artifactsDir, "CombinedImages.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the PDF file.");
    }

    private static void CreateSamplePng(string path)
    {
        // Create a 200x200 PNG with a light‑blue background and a dark‑blue rectangle.
        using (Bitmap bitmap = new Bitmap(200, 200, PixelFormat.Format32bppArgb))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
                using (Pen pen = new Pen(Color.DarkBlue, 5))
                {
                    graphics.DrawRectangle(pen, 20, 20, 160, 160);
                }
            }
            bitmap.Save(path, ImageFormat.Png);
        }
    }

    private static void CreateSampleJpeg(string path)
    {
        // Create a 200x200 JPEG with a light‑coral background and the word "JPEG".
        using (Bitmap bitmap = new Bitmap(200, 200, PixelFormat.Format24bppRgb))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightCoral);
                Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24);
                using (Brush brush = new SolidBrush(Color.White))
                {
                    graphics.DrawString("JPEG", font, brush, new PointF(40, 80));
                }
            }
            bitmap.Save(path, ImageFormat.Jpeg);
        }
    }
}

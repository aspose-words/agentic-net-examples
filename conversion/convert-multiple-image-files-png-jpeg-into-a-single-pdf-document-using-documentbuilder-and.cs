using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Prepare sample image files.
        string[] imagePaths = new string[] { "image1.png", "image2.jpg" };
        CreateSamplePng(imagePaths[0]);
        CreateSampleJpeg(imagePaths[1]);

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each image into the document, adding a page break between them.
        for (int i = 0; i < imagePaths.Length; i++)
        {
            builder.InsertImage(imagePaths[i]);
            if (i < imagePaths.Length - 1)
            {
                builder.InsertBreak(BreakType.PageBreak);
            }
        }

        // Save the document as a PDF file.
        string outputPath = "output.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }
    }

    private static void CreateSamplePng(string path)
    {
        // Create a 200x200 PNG image with some text.
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20))
                {
                    using (Brush brush = new SolidBrush(Color.Black))
                    {
                        graphics.DrawString("PNG Image", font, brush, new PointF(10, 80));
                    }
                }
            }
            bitmap.Save(path, ImageFormat.Png);
        }
    }

    private static void CreateSampleJpeg(string path)
    {
        // Create a 200x200 JPEG image with some text.
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightCoral);
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20))
                {
                    using (Brush brush = new SolidBrush(Color.White))
                    {
                        graphics.DrawString("JPEG Image", font, brush, new PointF(10, 80));
                    }
                }
            }
            bitmap.Save(path, ImageFormat.Jpeg);
        }
    }
}

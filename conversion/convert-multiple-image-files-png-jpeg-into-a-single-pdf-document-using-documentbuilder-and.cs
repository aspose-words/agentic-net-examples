using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary image files and the final PDF.
        const string pngPath = "sample_image.png";
        const string jpegPath = "sample_image.jpg";
        const string pdfPath = "combined_images.pdf";

        // Create sample PNG and JPEG images using Aspose.Drawing.
        CreateSamplePng(pngPath);
        CreateSampleJpeg(jpegPath);

        // Build a new Word document and insert the images.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the PNG image.
        builder.InsertImage(pngPath);
        builder.Writeln(); // Add a line break between images.

        // Insert the JPEG image.
        builder.InsertImage(jpegPath);

        // Save the document as a single PDF file.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Clean up temporary image files (optional).
        File.Delete(pngPath);
        File.Delete(jpegPath);
    }

    private static void CreateSamplePng(string filePath)
    {
        const int width = 300;
        const int height = 150;

        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);

                Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24);
                using (Brush brush = new SolidBrush(Color.Black))
                {
                    graphics.DrawString("PNG Sample", font, brush, new PointF(10, 60));
                }
                font.Dispose();
            }

            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    private static void CreateSampleJpeg(string filePath)
    {
        const int width = 300;
        const int height = 150;

        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightCoral);

                Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24);
                using (Brush brush = new SolidBrush(Color.White))
                {
                    graphics.DrawString("JPEG Sample", font, brush, new PointF(10, 60));
                }
                font.Dispose();
            }

            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }
}

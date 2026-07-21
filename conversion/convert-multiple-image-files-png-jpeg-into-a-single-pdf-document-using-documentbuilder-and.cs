using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;                 // Aspose.Drawing types for bitmap, graphics, colors, etc.
using Aspose.Drawing.Imaging;        // ImageFormat, PixelFormat

public class Program
{
    public static void Main()
    {
        // Paths for temporary images and final PDF.
        const string pngPath = "sample_image.png";
        const string jpegPath = "sample_image.jpg";
        const string outputPdf = "combined_images.pdf";

        // Create a PNG image.
        using (Bitmap pngBitmap = new Bitmap(200, 200, PixelFormat.Format24bppRgb))
        {
            using (Graphics graphics = Graphics.FromImage(pngBitmap))
            {
                graphics.Clear(Color.LightBlue);
                using (SolidBrush brush = new SolidBrush(Color.DarkBlue))
                {
                    // Use fully qualified Aspose.Drawing.Font to avoid ambiguity with Aspose.Words.Font.
                    using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
                    {
                        graphics.DrawString("PNG", font, brush, new PointF(50, 80));
                    }
                }
            }
            pngBitmap.Save(pngPath, ImageFormat.Png);
        }

        // Create a JPEG image.
        using (Bitmap jpegBitmap = new Bitmap(200, 200, PixelFormat.Format24bppRgb))
        {
            using (Graphics graphics = Graphics.FromImage(jpegBitmap))
            {
                graphics.Clear(Color.LightCoral);
                using (SolidBrush brush = new SolidBrush(Color.White))
                {
                    using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
                    {
                        graphics.DrawString("JPG", font, brush, new PointF(50, 80));
                    }
                }
            }
            jpegBitmap.Save(jpegPath, ImageFormat.Jpeg);
        }

        // Build a Word document and insert the images.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.InsertImage(pngPath);
        builder.InsertParagraph(); // Add spacing between images.
        builder.InsertImage(jpegPath);

        // Save the document as a PDF.
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("The PDF file was not created.");

        // Clean up temporary image files.
        File.Delete(pngPath);
        File.Delete(jpegPath);
    }
}

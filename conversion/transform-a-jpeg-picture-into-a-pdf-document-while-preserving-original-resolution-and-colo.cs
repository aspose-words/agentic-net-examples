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
        // Paths for the temporary JPEG image and the resulting PDF.
        const string imagePath = "sample.jpg";
        const string pdfPath = "image.pdf";

        // Create a sample JPEG image using Aspose.Drawing.
        CreateSampleJpeg(imagePath);

        // Create a new Word document and insert the JPEG image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);

        // Configure PDF save options to preserve the original image resolution and color depth.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ImageCompression = PdfImageCompression.Jpeg, // Keep JPEG format.
            JpegQuality = 100,                           // No quality loss.
            ColorMode = ColorMode.Normal                 // Preserve original colors.
        };
        // Disable downsampling to keep the original resolution.
        pdfOptions.DownsampleOptions.DownsampleImages = false;

        // Save the document as a PDF file.
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF was created successfully.
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("The PDF file was not created or is empty.");
    }

    private static void CreateSampleJpeg(string path)
    {
        // Create a 200x200 pixel bitmap with 24‑bit color depth.
        using (Bitmap bitmap = new Bitmap(200, 200, PixelFormat.Format24bppRgb))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with a light color.
                graphics.Clear(Color.LightBlue);

                // Draw a dark red ellipse in the center.
                using (SolidBrush brush = new SolidBrush(Color.DarkRed))
                {
                    graphics.FillEllipse(brush, 50, 50, 100, 100);
                }
            }

            // Save the bitmap as a JPEG image.
            bitmap.Save(path, ImageFormat.Jpeg);
        }
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define paths for the sample image and the resulting PDF.
        const string imagePath = "sample.jpg";
        const string pdfPath = "output.pdf";

        // Create a sample JPEG image using Aspose.Drawing.
        const int width = 800;
        const int height = 600;

        // Bitmap with 24‑bit RGB format.
        using (Bitmap bitmap = new Bitmap(width, height, PixelFormat.Format24bppRgb))
        {
            // Draw onto the bitmap.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with a solid color.
                graphics.Clear(Color.CornflowerBlue);

                // Draw a simple ellipse.
                using (Pen pen = new Pen(Color.Yellow, 5))
                {
                    graphics.DrawEllipse(pen, 100, 100, 600, 400);
                }
            }

            // Save the bitmap as a JPEG file, preserving resolution and color depth.
            bitmap.Save(imagePath, ImageFormat.Jpeg);
        }

        // Verify that the JPEG image was created.
        if (!File.Exists(imagePath))
            throw new InvalidOperationException("JPEG image was not created.");

        // Create a new Word document and insert the JPEG image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);

        // Save the document as a PDF file.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");
    }
}

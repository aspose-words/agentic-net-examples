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
        // Define file paths.
        string jpegPath = "sample.jpg";
        string pdfPath = "output.pdf";

        // Create a sample JPEG image using Aspose.Drawing.
        // The image is 800x600 pixels, 24‑bit color depth.
        Bitmap bitmap = new Bitmap(800, 600, PixelFormat.Format24bppRgb);
        try
        {
            Graphics graphics = Graphics.FromImage(bitmap);
            try
            {
                // Fill background with a solid color.
                graphics.Clear(Color.CornflowerBlue);

                // Draw a simple rectangle.
                using (Pen pen = new Pen(Color.Yellow, 5))
                {
                    graphics.DrawRectangle(pen, 100, 100, 600, 400);
                }
            }
            finally
            {
                graphics.Dispose();
            }

            // Save the bitmap as a JPEG file.
            bitmap.Save(jpegPath, ImageFormat.Jpeg);
        }
        finally
        {
            bitmap.Dispose();
        }

        // Verify that the JPEG file was created.
        if (!File.Exists(jpegPath))
            throw new InvalidOperationException("The JPEG image was not created.");

        // Create a new Word document and insert the JPEG image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(jpegPath);

        // Save the document as a PDF, preserving the original image resolution and color depth.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF document was not created.");

        // Optional: Clean up the temporary JPEG file.
        // File.Delete(jpegPath);
    }
}

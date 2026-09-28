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
        // Step 1: Create a sample JPEG image using Aspose.Drawing.
        const string jpegPath = "sample.jpg";
        const int imageWidth = 800;
        const int imageHeight = 600;

        using (Bitmap bitmap = new Bitmap(imageWidth, imageHeight, PixelFormat.Format24bppRgb))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with a solid color.
                graphics.Clear(Color.LightBlue);

                // Draw a simple ellipse.
                using (Pen pen = new Pen(Color.DarkBlue, 5))
                {
                    graphics.DrawEllipse(pen, 100, 100, 600, 400);
                }
            }

            // Save the bitmap as a JPEG file.
            bitmap.Save(jpegPath, ImageFormat.Jpeg);
        }

        // Verify that the JPEG file was created.
        if (!File.Exists(jpegPath))
            throw new InvalidOperationException("The JPEG image was not created.");

        // Step 2: Create a new Word document and insert the JPEG image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(jpegPath);

        // Step 3: Save the document as a PDF, preserving the original image resolution and color depth.
        const string pdfPath = "output.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF document was not created.");

        // Cleanup: optionally delete the temporary JPEG file.
        // File.Delete(jpegPath);
    }
}

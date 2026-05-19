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
        // Create a sample JPEG image using Aspose.Drawing.
        const string jpegPath = "sample.jpg";
        const int imageWidth = 800;
        const int imageHeight = 600;

        using (Bitmap bitmap = new Bitmap(imageWidth, imageHeight))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with a solid color.
                graphics.Clear(Color.CornflowerBlue);
                // Draw a simple ellipse.
                using (Pen pen = new Pen(Color.White, 5))
                {
                    graphics.DrawEllipse(pen, 100, 100, 600, 400);
                }
            }

            // Save the bitmap as a JPEG with maximum quality to preserve original data.
            bitmap.Save(jpegPath, ImageFormat.Jpeg);
        }

        // Verify that the JPEG file was created.
        if (!File.Exists(jpegPath))
            throw new InvalidOperationException("Failed to create the sample JPEG image.");

        // Create a new Word document and insert the JPEG image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(jpegPath);

        // Configure PDF save options to keep the original image resolution and color depth.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Use JPEG compression with the highest quality (no additional compression).
            ImageCompression = PdfImageCompression.Jpeg,
            JpegQuality = 100,
            // Disable downsampling to preserve the original resolution.
            DownsampleOptions = { DownsampleImages = false }
        };

        const string pdfPath = "output.pdf";
        doc.Save(pdfPath, pdfOptions);

        // Validate that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF conversion failed; output file not found.");

        // Clean up temporary JPEG file.
        File.Delete(jpegPath);
    }
}

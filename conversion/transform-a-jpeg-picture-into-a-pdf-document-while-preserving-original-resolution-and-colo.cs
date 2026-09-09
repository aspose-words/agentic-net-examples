using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging; // For ImageFormat

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string jpegPath = "sample.jpg";
        const string pdfPath = "output.pdf";

        // ------------------------------------------------------------
        // 1. Create a sample JPEG image using Aspose.Drawing.
        // ------------------------------------------------------------
        const int imageWidth = 800;
        const int imageHeight = 600;

        // Create a bitmap with the desired size.
        using (Bitmap bitmap = new Bitmap(imageWidth, imageHeight))
        {
            // Obtain a graphics object to draw on the bitmap.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with a solid color.
                graphics.Clear(Color.CornflowerBlue);

                // Draw a simple rectangle.
                using (Pen pen = new Pen(Color.Yellow, 5))
                {
                    graphics.DrawRectangle(pen, 50, 50, imageWidth - 100, imageHeight - 100);
                }

                // Draw some text using a drawing font.
                Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 48);
                try
                {
                    using (SolidBrush brush = new SolidBrush(Color.White))
                    {
                        graphics.DrawString("Sample JPEG", font, brush, new PointF(100, imageHeight / 2 - 24));
                    }
                }
                finally
                {
                    font.Dispose();
                }
            }

            // Save the bitmap as a JPEG with maximum quality to preserve color depth.
            using (MemoryStream jpegStream = new MemoryStream())
            {
                // Aspose.Drawing saves JPEG with high quality by default.
                bitmap.Save(jpegStream, ImageFormat.Jpeg);
                File.WriteAllBytes(jpegPath, jpegStream.ToArray());
            }
        }

        // Verify that the JPEG file was created.
        if (!File.Exists(jpegPath) || new FileInfo(jpegPath).Length == 0)
            throw new InvalidOperationException("Failed to create the sample JPEG image.");

        // ------------------------------------------------------------
        // 2. Create a Word document and insert the JPEG image.
        // ------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image at its original size (preserving resolution).
        builder.InsertImage(jpegPath);

        // ------------------------------------------------------------
        // 3. Save the document as PDF while preserving the original image quality.
        // ------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Prevent downsampling of images.
            DownsampleOptions = { DownsampleImages = false },

            // Preserve JPEG quality (no additional compression).
            JpegQuality = 100,

            // Use automatic image compression to keep original bytes when possible.
            ImageCompression = PdfImageCompression.Auto
        };

        doc.Save(pdfPath, pdfOptions);

        // ------------------------------------------------------------
        // 4. Validate that the PDF was created successfully.
        // ------------------------------------------------------------
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("PDF conversion failed; output file was not created.");
    }
}

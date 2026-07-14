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
        // Define file names.
        const string jpegPath = "sample.jpg";
        const string pdfPath = "output.pdf";

        // ------------------------------------------------------------
        // 1. Create a sample JPEG image using Aspose.Drawing.
        // ------------------------------------------------------------
        using (Bitmap bitmap = new Bitmap(800, 600))
        {
            // Fill the bitmap with a solid color.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.CornflowerBlue);
                // Draw a simple ellipse to have some content.
                using (Pen pen = new Pen(Color.White, 5))
                {
                    graphics.DrawEllipse(pen, 100, 100, 600, 400);
                }
            }

            // Save the bitmap as a JPEG with maximum quality (100).
            // This ensures the source image has full resolution and color depth.
            bitmap.Save(jpegPath, ImageFormat.Jpeg);
        }

        // Verify that the JPEG file was created.
        if (!File.Exists(jpegPath) || new FileInfo(jpegPath).Length == 0)
            throw new InvalidOperationException("Failed to create the source JPEG image.");

        // ------------------------------------------------------------
        // 2. Load the JPEG into a Word document.
        // ------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(jpegPath);

        // ------------------------------------------------------------
        // 3. Configure PDF save options to preserve the original image.
        // ------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Keep colors unchanged.
            ColorMode = ColorMode.Normal,
            // Preserve JPEG quality (no recompression).
            ImageCompression = PdfImageCompression.Jpeg,
            JpegQuality = 100,
            // Disable downsampling of images.
            DownsampleOptions = { DownsampleImages = false }
        };

        // ------------------------------------------------------------
        // 4. Save the document as PDF.
        // ------------------------------------------------------------
        doc.Save(pdfPath, pdfOptions);

        // ------------------------------------------------------------
        // 5. Validate that the PDF was created.
        // ------------------------------------------------------------
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("PDF conversion failed; output file was not created.");

        // Cleanup: optional removal of temporary files.
        // File.Delete(jpegPath);
    }
}

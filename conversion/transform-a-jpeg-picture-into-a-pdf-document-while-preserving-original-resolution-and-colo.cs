using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class JpegToPdfConverter
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Path for the temporary JPEG image.
        string jpegPath = Path.Combine(outputDir, "sample.jpg");

        // Create a sample JPEG image using Aspose.Drawing.
        // The image size (800x600) and color depth are preserved.
        using (Bitmap bitmap = new Bitmap(800, 600))
        {
            // Obtain a Graphics object from the bitmap.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                graphics.Clear(Color.White);

                // Draw a simple rectangle to have visible content.
                using (Pen pen = new Pen(Color.Blue, 5))
                {
                    graphics.DrawRectangle(pen, 100, 100, 600, 400);
                }
            }

            // Save the bitmap as a JPEG file with maximum quality (no compression).
            bitmap.Save(jpegPath, ImageFormat.Jpeg);
        }

        // Verify that the JPEG file was created.
        if (!File.Exists(jpegPath) || new FileInfo(jpegPath).Length == 0)
            throw new InvalidOperationException("Failed to create the sample JPEG image.");

        // Create a new blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the JPEG image into the document.
        builder.InsertImage(jpegPath);

        // Configure PDF save options to preserve the original image resolution and color depth.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Use JPEG compression with the highest quality to avoid re‑encoding loss.
            ImageCompression = PdfImageCompression.Jpeg,
            JpegQuality = 100
        };
        // Disable downsampling so the original resolution is kept.
        pdfOptions.DownsampleOptions.DownsampleImages = false;

        // Path for the resulting PDF file.
        string pdfPath = Path.Combine(outputDir, "result.pdf");

        // Save the document as PDF using the configured options.
        doc.Save(pdfPath, pdfOptions);

        // Validate that the PDF was created successfully.
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("PDF conversion failed; output file is missing or empty.");

        // Optionally, clean up the temporary JPEG file.
        // File.Delete(jpegPath);
    }
}

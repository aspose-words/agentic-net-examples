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
        // Define paths for the temporary PNG image and the resulting PDF.
        string baseDir = Directory.GetCurrentDirectory();
        string imagePath = Path.Combine(baseDir, "sample.png");
        string pdfPath = Path.Combine(baseDir, "sample.pdf");

        // -------------------------------------------------
        // 1. Create a simple PNG image using Aspose.Drawing.
        // -------------------------------------------------
        const int width = 200;
        const int height = 200;

        // Create a bitmap with 32‑bit ARGB pixel format.
        using (Bitmap bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb))
        {
            // Obtain a graphics object to draw on the bitmap.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                graphics.Clear(Color.White);

                // Draw a blue ellipse.
                using (Pen pen = new Pen(Color.Blue, 5))
                {
                    graphics.DrawEllipse(pen, 10, 10, width - 20, height - 20);
                }
            }

            // Save the bitmap as a PNG file.
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // 2. Load the PNG into a Word document and save as PDF.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the created PNG image.
        builder.InsertImage(imagePath);

        // Save the document as PDF using the native Save method.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // -------------------------------------------------
        // 3. Validate that the PDF was created.
        // -------------------------------------------------
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
        {
            throw new InvalidOperationException("PDF conversion failed: output file is missing or empty.");
        }

        // Optional cleanup of the temporary PNG (keep PDF as result).
        try
        {
            File.Delete(imagePath);
        }
        catch
        {
            // Ignored – if deletion fails, it does not affect the main task.
        }
    }
}

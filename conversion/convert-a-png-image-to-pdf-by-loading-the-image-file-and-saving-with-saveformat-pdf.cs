using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string pngPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.png");
        string pdfPath = Path.Combine(Directory.GetCurrentDirectory(), "output.pdf");

        // -----------------------------------------------------------------
        // Create a simple PNG image using Aspose.Drawing (no System.Drawing usage).
        // -----------------------------------------------------------------
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            // Obtain a Graphics object for drawing on the bitmap.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with light blue.
                graphics.Clear(Color.LightBlue);

                // Draw a red ellipse.
                using (Pen pen = new Pen(Color.Red, 5))
                {
                    graphics.DrawEllipse(pen, 20, 20, 160, 160);
                }

                // Draw some text.
                Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20);
                using (SolidBrush brush = new SolidBrush(Color.DarkBlue))
                {
                    graphics.DrawString("Sample", font, brush, new PointF(40, 80));
                }
            }

            // Save the bitmap as PNG.
            bitmap.Save(pngPath, ImageFormat.Png);
        }

        // Verify that the PNG file was created.
        if (!File.Exists(pngPath))
            throw new InvalidOperationException("Failed to create the PNG image.");

        // -----------------------------------------------------------------
        // Load the PNG into a Word document and save as PDF.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(pngPath);

        // Save the document as PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF conversion failed; output file not found.");

        // Cleanup: optional removal of the temporary PNG file.
        // File.Delete(pngPath);
    }
}

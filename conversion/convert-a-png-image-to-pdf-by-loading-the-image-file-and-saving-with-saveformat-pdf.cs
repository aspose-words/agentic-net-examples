using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Paths for temporary PNG and final PDF.
        const string pngPath = "sample.png";
        const string pdfPath = "output.pdf";

        // Image dimensions.
        const int width = 200;
        const int height = 200;

        // Create a bitmap and draw simple content onto it.
        using (Bitmap bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb))
        {
            // Obtain a Graphics object from the bitmap (Aspose.Drawing version).
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background.
                graphics.Clear(Color.LightBlue);

                // Create a drawing font (fully qualified to avoid ambiguity).
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20))
                {
                    // Draw text onto the image.
                    graphics.DrawString(
                        "Hello",
                        font,
                        new SolidBrush(Color.Black),
                        new PointF(10, 80));
                }
            }

            // Save the bitmap as a PNG file.
            bitmap.Save(pngPath, ImageFormat.Png);
        }

        // Load the PNG image as a document and convert it to PDF.
        Document doc = new Document(pngPath);
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Clean up the temporary PNG file.
        if (File.Exists(pngPath))
            File.Delete(pngPath);
    }
}

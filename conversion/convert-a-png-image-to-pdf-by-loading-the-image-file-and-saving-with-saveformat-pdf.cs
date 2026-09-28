using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a sample PNG image using Aspose.Drawing.
        string imagePath = "sample.png";
        int width = 200;
        int height = 100;

        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background.
                graphics.Clear(Color.LightBlue);

                // Prepare drawing font.
                Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 16);
                try
                {
                    // Draw text onto the image.
                    using (SolidBrush brush = new SolidBrush(Color.Black))
                    {
                        graphics.DrawString("Hello PNG", font, brush, new PointF(10, 40));
                    }
                }
                finally
                {
                    font.Dispose();
                }
            }

            // Save the bitmap as a PNG file.
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Load the PNG image into an Aspose.Words Document.
        Document doc = new Document(imagePath);

        // Convert and save the document as PDF.
        string pdfPath = "output.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Expected output PDF was not created.");
    }
}

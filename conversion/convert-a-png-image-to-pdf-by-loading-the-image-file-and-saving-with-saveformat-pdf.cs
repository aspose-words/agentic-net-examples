using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Paths for the input PNG and output PDF
        string imagePath = "sample.png";
        string pdfPath = "output.pdf";

        // Create a sample PNG image using Aspose.Drawing
        using (Bitmap bitmap = new Bitmap(400, 200))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with a light gray color
                graphics.Clear(Color.LightGray);

                // Prepare a drawing font
                Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24);
                try
                {
                    // Draw text onto the image
                    using (SolidBrush brush = new SolidBrush(Color.DarkBlue))
                    {
                        graphics.DrawString("Sample PNG", font, brush, new PointF(50, 80));
                    }
                }
                finally
                {
                    font.Dispose();
                }
            }

            // Save the bitmap as a PNG file
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Load the PNG image into an Aspose.Words Document
        Document doc = new Document(imagePath);

        // Save the document as PDF
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }
    }
}

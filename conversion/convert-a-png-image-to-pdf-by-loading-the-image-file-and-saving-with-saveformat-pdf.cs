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
        // Paths for the temporary PNG image and the resulting PDF.
        const string pngPath = "sample.png";
        const string pdfPath = "output.pdf";

        // Create a simple PNG image using Aspose.Drawing types.
        using (Bitmap bitmap = new Bitmap(300, 150))
        {
            // Obtain a Graphics object that can draw on the bitmap.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background.
                graphics.Clear(Color.LightGray);

                // Create a drawing font.
                Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24);
                try
                {
                    // Use a brush for the text color.
                    using (SolidBrush brush = new SolidBrush(Color.DarkBlue))
                    {
                        // Draw the string at the specified location.
                        graphics.DrawString("Sample PNG", font, brush, new PointF(20, 60));
                    }
                }
                finally
                {
                    font.Dispose();
                }
            }

            // Save the bitmap as a PNG file.
            bitmap.Save(pngPath, ImageFormat.Png);
        }

        // Load the PNG image as a document. Aspose.Words treats the image as a single‑page document.
        Document doc = new Document(pngPath);

        // Convert the document (image) to PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Clean up temporary PNG file (optional).
        File.Delete(pngPath);
    }
}

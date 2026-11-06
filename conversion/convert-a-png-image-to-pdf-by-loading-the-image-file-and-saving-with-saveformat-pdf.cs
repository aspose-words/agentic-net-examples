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
        // Define file names
        string pngPath = "sample.png";
        string pdfPath = "output.pdf";

        // Create a sample PNG image using Aspose.Drawing
        using (Bitmap bitmap = new Bitmap(200, 200, PixelFormat.Format32bppArgb))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background
                graphics.Clear(Color.LightBlue);
                // Draw a simple ellipse
                using (Pen pen = new Pen(Color.DarkBlue, 3))
                {
                    graphics.DrawEllipse(pen, 20, 20, 160, 160);
                }
            }

            // Save the bitmap as PNG
            bitmap.Save(pngPath, ImageFormat.Png);
        }

        // Verify the PNG was created
        if (!File.Exists(pngPath))
            throw new InvalidOperationException("The PNG image was not created.");

        // Load the PNG image into an Aspose.Words Document
        Document doc = new Document(pngPath);

        // Save the document as PDF
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Optionally, clean up the sample PNG (comment out if you want to keep it)
        // File.Delete(pngPath);
    }
}

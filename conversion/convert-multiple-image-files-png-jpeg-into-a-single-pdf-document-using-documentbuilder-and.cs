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
        // Set up working directories.
        string workingDir = Directory.GetCurrentDirectory();
        string imagesDir = Path.Combine(workingDir, "Images");
        Directory.CreateDirectory(imagesDir);

        string pngPath = Path.Combine(imagesDir, "sample.png");
        string jpegPath = Path.Combine(imagesDir, "sample.jpg");
        string pdfPath = Path.Combine(workingDir, "CombinedImages.pdf");

        // Create a simple PNG image using Aspose.Drawing.
        using (Bitmap pngBitmap = new Bitmap(200, 200))
        {
            using (Graphics graphics = Graphics.FromImage(pngBitmap))
            {
                graphics.Clear(Color.LightBlue);
                // Use fully qualified Aspose.Drawing.Font to avoid ambiguity.
                Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20);
                graphics.DrawString("PNG Image", font, new SolidBrush(Color.DarkBlue), new PointF(20, 80));
                font.Dispose();
            }
            pngBitmap.Save(pngPath, ImageFormat.Png);
        }

        // Create a simple JPEG image using Aspose.Drawing.
        using (Bitmap jpegBitmap = new Bitmap(200, 200))
        {
            using (Graphics graphics = Graphics.FromImage(jpegBitmap))
            {
                graphics.Clear(Color.LightCoral);
                Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20);
                graphics.DrawString("JPEG Image", font, new SolidBrush(Color.White), new PointF(20, 80));
                font.Dispose();
            }
            jpegBitmap.Save(jpegPath, ImageFormat.Jpeg);
        }

        // Build a Word document and insert the images.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.InsertImage(pngPath);
        builder.InsertBreak(BreakType.PageBreak);
        builder.InsertImage(jpegPath);

        // Save the document as a single PDF file.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Optional cleanup of temporary images.
        try
        {
            File.Delete(pngPath);
            File.Delete(jpegPath);
            Directory.Delete(imagesDir);
        }
        catch
        {
            // Ignored – cleanup is best‑effort.
        }
    }
}

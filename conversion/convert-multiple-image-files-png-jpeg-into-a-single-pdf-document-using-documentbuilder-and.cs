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
        // Prepare a folder for temporary image files.
        string imagesFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputImages");
        Directory.CreateDirectory(imagesFolder);

        // Create a PNG image.
        string pngPath = Path.Combine(imagesFolder, "sample.png");
        using (Bitmap pngBitmap = new Bitmap(200, 200))
        {
            using (Graphics graphics = Graphics.FromImage(pngBitmap))
            {
                graphics.Clear(Color.LightBlue);
                using (Pen pen = new Pen(Color.Red, 5))
                {
                    graphics.DrawEllipse(pen, 20, 20, 160, 160);
                }
            }
            pngBitmap.Save(pngPath, ImageFormat.Png);
        }

        // Create a JPEG image.
        string jpegPath = Path.Combine(imagesFolder, "sample.jpg");
        using (Bitmap jpegBitmap = new Bitmap(200, 200))
        {
            using (Graphics graphics = Graphics.FromImage(jpegBitmap))
            {
                graphics.Clear(Color.LightGreen);
                using (Pen pen = new Pen(Color.Blue, 5))
                {
                    graphics.DrawRectangle(pen, 30, 30, 140, 140);
                }
            }
            jpegBitmap.Save(jpegPath, ImageFormat.Jpeg);
        }

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the PNG image.
        builder.InsertImage(pngPath);
        // Insert a page break between images.
        builder.InsertBreak(BreakType.PageBreak);
        // Insert the JPEG image.
        builder.InsertImage(jpegPath);

        // Save the document as a PDF.
        string outputPdf = Path.Combine(Directory.GetCurrentDirectory(), "ImagesCombined.pdf");
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("The PDF file was not created.");

        // Clean up temporary images (optional).
        try
        {
            File.Delete(pngPath);
            File.Delete(jpegPath);
            Directory.Delete(imagesFolder);
        }
        catch
        {
            // Ignored – cleanup is not critical for the example.
        }
    }
}

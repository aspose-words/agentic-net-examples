using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class ImageToPdfConverter
{
    public static void Main()
    {
        // Prepare a temporary folder for sample images.
        string imagesFolder = Path.Combine(Directory.GetCurrentDirectory(), "Images");
        Directory.CreateDirectory(imagesFolder);

        // Create a sample PNG image.
        string pngPath = Path.Combine(imagesFolder, "sample1.png");
        CreateSampleImage(pngPath, ImageFormat.Png, Color.LightBlue, Color.Red);

        // Create a sample JPEG image.
        string jpegPath = Path.Combine(imagesFolder, "sample2.jpg");
        CreateSampleImage(jpegPath, ImageFormat.Jpeg, Color.LightGreen, Color.Blue);

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the PNG image.
        builder.InsertImage(pngPath);
        // Insert a page break before the next image.
        builder.InsertBreak(BreakType.PageBreak);

        // Insert the JPEG image.
        builder.InsertImage(jpegPath);

        // Save the document as PDF.
        string outputPdf = Path.Combine(Directory.GetCurrentDirectory(), "output.pdf");
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("Expected output PDF was not created.");

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

    private static void CreateSampleImage(string filePath, ImageFormat format, Color backgroundColor, Color shapeColor)
    {
        const int width = 300;
        const int height = 200;

        using (Bitmap bitmap = new Bitmap(width, height))
        {
            // Obtain a Graphics object from the bitmap.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background.
                using (SolidBrush backgroundBrush = new SolidBrush(backgroundColor))
                {
                    graphics.FillRectangle(backgroundBrush, 0, 0, width, height);
                }

                // Draw a colored ellipse.
                using (SolidBrush shapeBrush = new SolidBrush(shapeColor))
                {
                    graphics.FillEllipse(shapeBrush, 50, 30, width - 100, height - 60);
                }
            }

            // Save the image in the requested format.
            bitmap.Save(filePath, format);
        }
    }
}

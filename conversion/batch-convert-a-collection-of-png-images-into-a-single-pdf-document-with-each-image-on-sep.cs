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
        // Create a folder to hold sample PNG images.
        string imagesFolder = "Images";
        Directory.CreateDirectory(imagesFolder);

        // Generate a few sample PNG images using Aspose.Drawing.
        for (int i = 1; i <= 3; i++)
        {
            string imagePath = Path.Combine(imagesFolder, $"image{i}.png");
            using (Bitmap bitmap = new Bitmap(200, 200))
            {
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    // Fill the bitmap with a distinct background color.
                    int red = 100 + i * 30;
                    graphics.Clear(Color.FromArgb(255, red, 150, 200));
                }
                // Save the bitmap as PNG.
                bitmap.Save(imagePath, ImageFormat.Png);
            }
        }

        // Create a new empty Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Get all generated PNG files.
        string[] pngFiles = Directory.GetFiles(imagesFolder, "*.png");
        bool firstImage = true;

        // Insert each PNG on a separate page.
        foreach (string pngFile in pngFiles)
        {
            if (!firstImage)
            {
                builder.InsertBreak(BreakType.PageBreak);
            }
            builder.InsertImage(pngFile);
            firstImage = false;
        }

        // Save the document as a single PDF file.
        string outputPdf = "output.pdf";
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPdf))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }
    }
}

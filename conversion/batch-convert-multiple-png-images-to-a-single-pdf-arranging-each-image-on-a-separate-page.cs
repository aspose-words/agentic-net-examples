using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Define folders and file names.
        string imagesFolder = "InputImages";
        string outputPdf = "CombinedImages.pdf";

        // Ensure the images folder exists.
        Directory.CreateDirectory(imagesFolder);

        // Create sample PNG images using Aspose.Drawing.
        for (int i = 1; i <= 3; i++)
        {
            string imagePath = Path.Combine(imagesFolder, $"Image{i}.png");
            CreateSamplePng(imagePath, i);
        }

        // Create a new blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each PNG image on a separate page.
        string[] pngFiles = Directory.GetFiles(imagesFolder, "*.png");
        for (int i = 0; i < pngFiles.Length; i++)
        {
            builder.InsertImage(pngFiles[i]);

            // Add a page break after each image except the last one.
            if (i < pngFiles.Length - 1)
                builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
        }

        // Save the document as a PDF.
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("The PDF file was not created.");

        // Optional cleanup.
        // Directory.Delete(imagesFolder, true);
    }

    // Creates a simple PNG image with a solid background and a number drawn on it.
    private static void CreateSamplePng(string filePath, int number)
    {
        // Define image size.
        int width = 200;
        int height = 200;

        // Create a bitmap and obtain a graphics object.
        using (Bitmap bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            // Fill background with a distinct color.
            graphics.Clear(Color.FromArgb(255, 100 + number * 30, 150, 200));

            // Prepare a drawing font.
            Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 48, FontStyle.Bold);
            try
            {
                // Draw the image number centered.
                string text = $"Img {number}";
                SizeF textSize = graphics.MeasureString(text, font);
                PointF location = new PointF((width - textSize.Width) / 2, (height - textSize.Height) / 2);
                graphics.DrawString(text, font, Brushes.White, location);
            }
            finally
            {
                font.Dispose();
            }

            // Save the bitmap as PNG.
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}

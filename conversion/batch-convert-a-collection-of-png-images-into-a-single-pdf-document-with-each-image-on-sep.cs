using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;               // Aspose.Drawing types for bitmap, graphics, colors, fonts, etc.
using Aspose.Drawing.Imaging;      // For ImageFormat

public class Program
{
    public static void Main()
    {
        // Define folders and file names.
        string baseDir = Directory.GetCurrentDirectory();
        string imagesDir = Path.Combine(baseDir, "Images");
        string outputPdf = Path.Combine(baseDir, "CombinedImages.pdf");

        // Ensure the images folder exists.
        if (!Directory.Exists(imagesDir))
            Directory.CreateDirectory(imagesDir);

        // Create sample PNG images.
        const int imageCount = 3;
        for (int i = 1; i <= imageCount; i++)
        {
            string imagePath = Path.Combine(imagesDir, $"Image{i}.png");
            CreateSamplePng(imagePath, i);
        }

        // Create a new blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each PNG onto a separate page.
        for (int i = 1; i <= imageCount; i++)
        {
            string imagePath = Path.Combine(imagesDir, $"Image{i}.png");
            builder.InsertImage(imagePath);
            if (i < imageCount)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document as PDF.
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("The PDF file was not created.");

        // Optional: clean up generated files (comment out if you want to keep them).
        //Directory.Delete(imagesDir, true);
    }

    // Generates a simple PNG image with a solid background and a label.
    private static void CreateSamplePng(string filePath, int index)
    {
        const int width = 400;
        const int height = 300;

        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with a distinct color per image.
                Aspose.Drawing.Color bgColor = index switch
                {
                    1 => Aspose.Drawing.Color.LightCoral,
                    2 => Aspose.Drawing.Color.LightGreen,
                    _ => Aspose.Drawing.Color.LightBlue
                };
                graphics.Clear(bgColor);

                // Draw the image label.
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
                using (SolidBrush brush = new SolidBrush(Aspose.Drawing.Color.Black))
                {
                    string text = $"Sample Image {index}";
                    graphics.DrawString(text, font, brush, new PointF(20, height / 2 - 12));
                }
            }

            // Save as PNG.
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}

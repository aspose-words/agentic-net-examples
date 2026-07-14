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
        // Define folders and file names.
        string baseDir = Directory.GetCurrentDirectory();
        string imagesDir = Path.Combine(baseDir, "InputImages");
        string outputPdfPath = Path.Combine(baseDir, "ImagesCombined.pdf");

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

            // Add a page break after each image except the last one.
            if (i < imageCount)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document as a PDF.
        doc.Save(outputPdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException("The output PDF file was not created.");

        // Clean up temporary images (optional).
        // Directory.Delete(imagesDir, true);
    }

    // Generates a simple PNG image with a solid background and a number drawn on it.
    private static void CreateSamplePng(string filePath, int number)
    {
        const int width = 200;
        const int height = 200;

        // Create a bitmap using Aspose.Drawing.
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            // Fill the bitmap with a color that varies by number.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Choose a background color.
                Color backgroundColor = number % 3 == 1 ? Color.LightBlue :
                                        number % 3 == 2 ? Color.LightGreen :
                                        Color.LightCoral;
                graphics.Clear(backgroundColor);

                // Draw the number in the center using a simple font.
                // Note: Aspose.Drawing.Font is used to avoid ambiguity with Aspose.Words.Font.
                Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 48);
                try
                {
                    // Measure the string to center it.
                    SizeF textSize = graphics.MeasureString(number.ToString(), font);
                    float x = (width - textSize.Width) / 2;
                    float y = (height - textSize.Height) / 2;
                    graphics.DrawString(number.ToString(), font, Brushes.Black, x, y);
                }
                finally
                {
                    font.Dispose();
                }
            }

            // Save the bitmap as a PNG file.
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}

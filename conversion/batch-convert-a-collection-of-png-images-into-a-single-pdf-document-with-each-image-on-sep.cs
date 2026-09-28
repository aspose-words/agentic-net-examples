using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Create a folder for sample PNG images.
        string imagesFolder = "Images";
        Directory.CreateDirectory(imagesFolder);

        // Generate a few sample PNG files.
        int imageCount = 3;
        for (int i = 1; i <= imageCount; i++)
        {
            string imagePath = Path.Combine(imagesFolder, $"image{i}.png");
            CreateSamplePng(imagePath, i);
        }

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each PNG onto a separate page.
        string[] pngFiles = Directory.GetFiles(imagesFolder, "*.png");
        for (int i = 0; i < pngFiles.Length; i++)
        {
            builder.InsertImage(pngFiles[i]);

            // Add a page break after each image except the last one.
            if (i < pngFiles.Length - 1)
            {
                builder.InsertBreak(BreakType.PageBreak);
            }
        }

        // Save the document as a PDF file.
        string outputPdf = "output.pdf";
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPdf))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }
    }

    private static void CreateSamplePng(string path, int index)
    {
        int width = 600;
        int height = 800;

        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background.
                graphics.Clear(Color.LightBlue);

                // Draw a rectangle border.
                using (Pen pen = new Pen(Color.DarkBlue, 5))
                {
                    graphics.DrawRectangle(pen, 50, 50, width - 100, height - 100);
                }

                // Draw centered text.
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 48))
                {
                    using (Brush brush = new SolidBrush(Color.Black))
                    {
                        string text = $"Image {index}";
                        SizeF textSize = graphics.MeasureString(text, font);
                        float x = (width - textSize.Width) / 2;
                        float y = (height - textSize.Height) / 2;
                        graphics.DrawString(text, font, brush, x, y);
                    }
                }
            }

            // Save the bitmap as a PNG file.
            bitmap.Save(path, ImageFormat.Png);
        }
    }
}

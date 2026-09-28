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
        // Create a folder to hold sample PNG images.
        string imagesFolder = "Images";
        Directory.CreateDirectory(imagesFolder);

        // Generate sample PNG images using Aspose.Drawing.
        int imageCount = 3;
        string[] imagePaths = new string[imageCount];
        for (int i = 0; i < imageCount; i++)
        {
            string filePath = Path.Combine(imagesFolder, $"Image{i + 1}.png");
            CreateSamplePng(filePath, i + 1);
            imagePaths[i] = filePath;
        }

        // Create a new empty Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each PNG on a separate page.
        for (int i = 0; i < imagePaths.Length; i++)
        {
            if (i > 0)
            {
                // Insert a page break before adding the next image.
                builder.InsertBreak(BreakType.PageBreak);
            }

            // Insert the image. The image will be scaled to fit the page width.
            builder.InsertImage(imagePaths[i]);
        }

        // Save the document as a single PDF file.
        string outputPdf = "CombinedImages.pdf";
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPdf))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }

        // Clean up generated images (optional).
        // Directory.Delete(imagesFolder, true);
    }

    private static void CreateSamplePng(string filePath, int index)
    {
        const int width = 400;
        const int height = 300;

        // Create a bitmap with a 32bpp ARGB pixel format.
        using (Bitmap bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb))
        {
            // Obtain a graphics object for drawing.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background.
                using (SolidBrush backgroundBrush = new SolidBrush(Color.White))
                {
                    graphics.FillRectangle(backgroundBrush, new Rectangle(0, 0, width, height));
                }

                // Draw a colored rectangle.
                using (Pen pen = new Pen(Color.Blue, 5))
                {
                    Rectangle rect = new Rectangle(50, 50, width - 100, height - 100);
                    graphics.DrawRectangle(pen, rect);
                }

                // Draw sample text.
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
                {
                    using (SolidBrush textBrush = new SolidBrush(Color.Black))
                    {
                        string text = $"Sample Image {index}";
                        SizeF textSize = graphics.MeasureString(text, font);
                        PointF textPosition = new PointF((width - textSize.Width) / 2, (height - textSize.Height) / 2);
                        graphics.DrawString(text, font, textBrush, textPosition);
                    }
                }
            }

            // Save the bitmap as a PNG file.
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}

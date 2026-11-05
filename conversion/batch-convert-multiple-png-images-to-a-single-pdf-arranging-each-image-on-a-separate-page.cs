using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;
using Aspose.Drawing.Text;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for sample PNG images.
        string imagesFolder = Path.Combine(Directory.GetCurrentDirectory(), "Images");
        Directory.CreateDirectory(imagesFolder);

        // Create three sample PNG images using Aspose.Drawing.
        for (int i = 1; i <= 3; i++)
        {
            string imagePath = Path.Combine(imagesFolder, $"image{i}.png");
            using (Bitmap bitmap = new Bitmap(400, 300, PixelFormat.Format32bppArgb))
            {
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    // Fill background.
                    using (SolidBrush backgroundBrush = new SolidBrush(Color.LightBlue))
                    {
                        graphics.FillRectangle(backgroundBrush, 0, 0, bitmap.Width, bitmap.Height);
                    }

                    // Draw text.
                    using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
                    {
                        using (SolidBrush textBrush = new SolidBrush(Color.Black))
                        {
                            string text = $"Image {i}";
                            SizeF textSize = graphics.MeasureString(text, font);
                            PointF location = new PointF((bitmap.Width - textSize.Width) / 2, (bitmap.Height - textSize.Height) / 2);
                            graphics.DrawString(text, font, textBrush, location);
                        }
                    }
                }

                // Save as PNG.
                bitmap.Save(imagePath, ImageFormat.Png);
            }

            if (!File.Exists(imagePath))
                throw new InvalidOperationException($"Failed to create sample image: {imagePath}");
        }

        // Create a new empty Word document.
        Document pdfDocument = new Document();
        DocumentBuilder builder = new DocumentBuilder(pdfDocument);

        // Insert each PNG image on a separate page.
        string[] pngFiles = Directory.GetFiles(imagesFolder, "*.png");
        for (int index = 0; index < pngFiles.Length; index++)
        {
            if (index > 0)
                builder.InsertBreak(BreakType.PageBreak);

            // Center the image on the page.
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            // Insert the image.
            builder.InsertImage(pngFiles[index]);
        }

        // Save the document as a single PDF file.
        string outputPdfPath = Path.Combine(Directory.GetCurrentDirectory(), "ImagesCombined.pdf");
        pdfDocument.Save(outputPdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException("The combined PDF file was not created.");

        // Optionally, clean up the temporary images folder.
        // Directory.Delete(imagesFolder, true);
    }
}

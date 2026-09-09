using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Folder to hold generated PNG images.
        const string imagesFolder = "InputImages";
        Directory.CreateDirectory(imagesFolder);

        // Number of sample images to create.
        const int imageCount = 3;

        // Create sample PNG images using Aspose.Drawing.
        for (int i = 1; i <= imageCount; i++)
        {
            string filePath = Path.Combine(imagesFolder, $"Image{i}.png");
            using (Bitmap bitmap = new Bitmap(600, 800, PixelFormat.Format32bppArgb))
            {
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    // Fill background.
                    graphics.Clear(Color.LightBlue);

                    // Prepare drawing objects.
                    Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 36);
                    using (SolidBrush brush = new SolidBrush(Color.DarkBlue))
                    {
                        string text = $"Sample Image {i}";
                        graphics.DrawString(text, font, brush, new PointF(50, 350));
                    }
                }

                // Save as PNG.
                bitmap.Save(filePath, ImageFormat.Png);
            }
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
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the assembled document as PDF.
        const string outputPdf = "CombinedImages.pdf";
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("The PDF file was not created as expected.");

        // Cleanup temporary images (optional).
        foreach (string file in pngFiles)
        {
            File.Delete(file);
        }
        Directory.Delete(imagesFolder);
    }
}

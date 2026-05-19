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
        // Prepare a folder for the sample PNG images.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputImages");
        Directory.CreateDirectory(inputFolder);

        // Create three sample PNG images using Aspose.Drawing.
        for (int i = 1; i <= 3; i++)
        {
            string imagePath = Path.Combine(inputFolder, $"image{i}.png");
            using (Bitmap bitmap = new Bitmap(300, 300))
            {
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    // Fill background.
                    graphics.Clear(Color.White);

                    // Draw a colored ellipse.
                    using (SolidBrush ellipseBrush = new SolidBrush(Color.FromArgb(100, 150, 200)))
                    {
                        graphics.FillEllipse(ellipseBrush, 30, 30, 240, 240);
                    }

                    // Draw image label text.
                    using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
                    using (SolidBrush textBrush = new SolidBrush(Color.Black))
                    {
                        graphics.DrawString($"Image {i}", font, textBrush, new PointF(80, 130));
                    }
                }

                // Save the bitmap as PNG.
                bitmap.Save(imagePath, ImageFormat.Png);
            }
        }

        // Create a new blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each PNG image on a separate page.
        string[] pngFiles = Directory.GetFiles(inputFolder, "*.png");
        Array.Sort(pngFiles); // Ensure deterministic order.

        for (int i = 0; i < pngFiles.Length; i++)
        {
            builder.InsertImage(pngFiles[i]);

            // Add a page break after each image except the last one.
            if (i < pngFiles.Length - 1)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the assembled document as a PDF.
        string outputPdf = Path.Combine(Directory.GetCurrentDirectory(), "output.pdf");
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("The PDF file was not created as expected.");
    }
}

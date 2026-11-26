using System;
using System.Collections.Generic;
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
        // Prepare a temporary folder for sample images.
        string tempFolder = Path.Combine(Directory.GetCurrentDirectory(), "TempImages");
        Directory.CreateDirectory(tempFolder);

        // Create sample PNG images.
        List<string> imagePaths = new List<string>();
        for (int i = 1; i <= 3; i++)
        {
            string imagePath = Path.Combine(tempFolder, $"image{i}.png");
            CreateSamplePng(imagePath, i);
            imagePaths.Add(imagePath);
        }

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each image on a separate page.
        for (int i = 0; i < imagePaths.Count; i++)
        {
            builder.InsertImage(imagePaths[i]);

            // Add a page break after each image except the last one.
            if (i < imagePaths.Count - 1)
            {
                builder.InsertBreak(BreakType.PageBreak);
            }
        }

        // Save the document as PDF.
        string outputPdf = Path.Combine(Directory.GetCurrentDirectory(), "ImagesCombined.pdf");
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPdf))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }

        // Clean up temporary images.
        foreach (string path in imagePaths)
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }

        // Optionally delete the temporary folder if empty.
        if (Directory.Exists(tempFolder) && Directory.GetFiles(tempFolder).Length == 0)
        {
            Directory.Delete(tempFolder);
        }
    }

    private static void CreateSamplePng(string filePath, int index)
    {
        // Create a 200x200 bitmap.
        using (Bitmap bitmap = new Bitmap(200, 200, PixelFormat.Format32bppArgb))
        {
            // Draw on the bitmap.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with a color that varies by index.
                Color backgroundColor = index % 2 == 0 ? Color.LightGreen : Color.LightBlue;
                using (SolidBrush brush = new SolidBrush(backgroundColor))
                {
                    graphics.FillRectangle(brush, 0, 0, bitmap.Width, bitmap.Height);
                }

                // Draw a simple ellipse.
                using (Pen pen = new Pen(Color.DarkGray, 3))
                {
                    graphics.DrawEllipse(pen, 20, 20, bitmap.Width - 40, bitmap.Height - 40);
                }

                // Draw index text.
                Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24);
                using (SolidBrush textBrush = new SolidBrush(Color.Black))
                {
                    StringFormat format = new StringFormat
                    {
                        Alignment = StringAlignment.Center,
                        LineAlignment = StringAlignment.Center
                    };
                    graphics.DrawString($"Image {index}", font, textBrush, new RectangleF(0, 0, bitmap.Width, bitmap.Height), format);
                }
            }

            // Save the bitmap as PNG.
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}

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
        // Directory to hold the generated PNG images.
        string imagesDir = "InputImages";
        Directory.CreateDirectory(imagesDir);

        // Generate a few sample PNG files.
        for (int i = 0; i < 3; i++)
        {
            string imagePath = Path.Combine(imagesDir, $"image{i}.png");
            CreateSamplePng(imagePath, i);
        }

        // Create a new blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each PNG onto its own page.
        string[] pngFiles = Directory.GetFiles(imagesDir, "*.png");
        for (int i = 0; i < pngFiles.Length; i++)
        {
            builder.InsertImage(pngFiles[i]);

            // Add a page break after every image except the last one.
            if (i < pngFiles.Length - 1)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the assembled document as a PDF.
        string outputPdf = "output.pdf";
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("The PDF file was not created.");

        // Optional: clean up the temporary PNG files.
        // foreach (string file in pngFiles) File.Delete(file);
    }

    // Creates a simple PNG image with a colored background and some text.
    private static void CreateSamplePng(string path, int index)
    {
        // 200x200 pixel bitmap.
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Alternate background colors.
                Color background = (index % 2 == 0) ? Color.LightBlue : Color.LightGreen;
                graphics.Clear(background);

                // Draw identifying text.
                Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20);
                using (Brush brush = new SolidBrush(Color.Black))
                {
                    string text = $"Image {index + 1}";
                    graphics.DrawString(text, font, brush, new PointF(10, 80));
                }

                // Save the bitmap as a PNG file.
                bitmap.Save(path, ImageFormat.Png);
            }
        }
    }
}

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
        // Define folders for temporary images and final PDF.
        string baseDir = Directory.GetCurrentDirectory();
        string imagesDir = Path.Combine(baseDir, "InputImages");
        string outputPdfPath = Path.Combine(baseDir, "ImagesCombined.pdf");

        // Ensure the images folder exists.
        Directory.CreateDirectory(imagesDir);

        // Create sample PNG images using Aspose.Drawing.
        string[] imagePaths = new string[3];
        for (int i = 0; i < imagePaths.Length; i++)
        {
            string filePath = Path.Combine(imagesDir, $"Sample{i + 1}.png");
            CreateSamplePng(filePath, i);
            imagePaths[i] = filePath;
        }

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each image on a separate page.
        for (int i = 0; i < imagePaths.Length; i++)
        {
            builder.InsertImage(imagePaths[i]);

            // Add a page break after each image except the last one.
            if (i < imagePaths.Length - 1)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document as PDF.
        doc.Save(outputPdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(outputPdfPath))
            throw new FileNotFoundException("The PDF output file was not created.", outputPdfPath);
    }

    // Generates a simple PNG image with a solid background and some text.
    private static void CreateSamplePng(string filePath, int index)
    {
        // Define image size.
        int width = 600;
        int height = 400;

        // Choose a background color based on the index.
        Color backgroundColor = index switch
        {
            0 => Color.LightBlue,
            1 => Color.LightGreen,
            _ => Color.LightCoral
        };

        // Create bitmap and graphics objects.
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            // Obtain a Graphics object that can draw onto the bitmap.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background.
                graphics.Clear(backgroundColor);

                // Draw sample text.
                Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24);
                using (SolidBrush brush = new SolidBrush(Color.Black))
                {
                    string text = $"Sample Image {index + 1}";
                    graphics.DrawString(text, font, brush, new PointF(50, 180));
                }
            }

            // Save as PNG.
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}

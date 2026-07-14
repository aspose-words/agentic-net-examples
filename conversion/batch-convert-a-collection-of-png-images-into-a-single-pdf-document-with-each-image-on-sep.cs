using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Folder for temporary image files.
        string workDir = Directory.GetCurrentDirectory();

        // Create sample PNG images using Aspose.Drawing.
        string[] imageFiles = new string[3];
        for (int i = 0; i < imageFiles.Length; i++)
        {
            string filePath = Path.Combine(workDir, $"sample{i + 1}.png");
            CreateSamplePng(filePath, $"Image {i + 1}");
            imageFiles[i] = filePath;
        }

        // Build a Word document where each image occupies a separate page.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 0; i < imageFiles.Length; i++)
        {
            builder.InsertImage(imageFiles[i]);

            // Insert a page break after each image except the last one.
            if (i < imageFiles.Length - 1)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document as a single PDF file.
        string pdfPath = Path.Combine(workDir, "ImagesCombined.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Clean up temporary PNG files.
        foreach (string file in imageFiles)
        {
            if (File.Exists(file))
                File.Delete(file);
        }
    }

    // Generates a simple PNG image with centered text using Aspose.Drawing.
    private static void CreateSamplePng(string filePath, string text)
    {
        const int width = 400;
        const int height = 300;

        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background.
                graphics.Clear(Color.LightGray);

                // Define font and brush.
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
                using (SolidBrush brush = new SolidBrush(Color.Black))
                {
                    // Measure text size.
                    SizeF textSize = graphics.MeasureString(text, font);
                    // Calculate position to center the text.
                    float x = (width - textSize.Width) / 2;
                    float y = (height - textSize.Height) / 2;

                    // Draw the text.
                    graphics.DrawString(text, font, brush, new PointF(x, y));
                }
            }

            // Save as PNG.
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}

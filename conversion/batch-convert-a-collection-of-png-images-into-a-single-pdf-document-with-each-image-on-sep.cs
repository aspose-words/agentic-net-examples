using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing; // Aspose.Drawing.Common namespace

public class BatchPngToPdf
{
    public static void Main()
    {
        // Define folders for input images and output PDF.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputImages");
        string outputPdfPath = Path.Combine(Directory.GetCurrentDirectory(), "CombinedImages.pdf");

        // Ensure the input folder exists.
        Directory.CreateDirectory(inputFolder);

        // Create sample PNG images using Aspose.Drawing.
        CreateSamplePng(Path.Combine(inputFolder, "Image1.png"), Aspose.Drawing.Color.LightBlue, "First");
        CreateSamplePng(Path.Combine(inputFolder, "Image2.png"), Aspose.Drawing.Color.LightGreen, "Second");
        CreateSamplePng(Path.Combine(inputFolder, "Image3.png"), Aspose.Drawing.Color.LightCoral, "Third");

        // Gather all PNG files from the input folder.
        string[] pngFiles = Directory.GetFiles(inputFolder, "*.png");

        if (pngFiles.Length == 0)
            throw new InvalidOperationException("No PNG images were found to convert.");

        // Create a new blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each PNG onto a separate page.
        for (int i = 0; i < pngFiles.Length; i++)
        {
            builder.InsertImage(pngFiles[i]);

            // Add a page break after each image except the last one.
            if (i < pngFiles.Length - 1)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the assembled document as a PDF.
        doc.Save(outputPdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        Console.WriteLine($"Successfully created PDF: {outputPdfPath}");
    }

    // Helper method to create a simple PNG image with a solid background and centered text.
    private static void CreateSamplePng(string filePath, Aspose.Drawing.Color backgroundColor, string label)
    {
        const int width = 400;
        const int height = 300;

        // Create a bitmap and obtain a graphics object for drawing.
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            // Fill background.
            using (SolidBrush brush = new SolidBrush(backgroundColor))
            {
                graphics.FillRectangle(brush, 0, 0, width, height);
            }

            // Prepare font and text layout.
            Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24);
            using (SolidBrush textBrush = new SolidBrush(Aspose.Drawing.Color.Black))
            {
                // Measure the text size.
                SizeF textSize = graphics.MeasureString(label, font);
                float x = (width - textSize.Width) / 2;
                float y = (height - textSize.Height) / 2;

                // Draw the label text.
                graphics.DrawString(label, font, textBrush, x, y);
            }

            // Save as PNG.
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }
    }
}

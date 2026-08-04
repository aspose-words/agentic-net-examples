using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;               // Aspose.Drawing.Common namespace
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Folder to hold generated PNG images
        string imagesFolder = "InputImages";
        Directory.CreateDirectory(imagesFolder);

        // Create three sample PNG images using Aspose.Drawing
        CreateSamplePng(Path.Combine(imagesFolder, "Image1.png"), Aspose.Drawing.Color.Red, "Image 1");
        CreateSamplePng(Path.Combine(imagesFolder, "Image2.png"), Aspose.Drawing.Color.Green, "Image 2");
        CreateSamplePng(Path.Combine(imagesFolder, "Image3.png"), Aspose.Drawing.Color.Blue, "Image 3");

        // Create a new Word document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Get all PNG files from the folder
        string[] pngFiles = Directory.GetFiles(imagesFolder, "*.png");

        for (int i = 0; i < pngFiles.Length; i++)
        {
            // Insert the image onto the current page
            builder.InsertImage(pngFiles[i]);

            // Add a page break after each image except the last one
            if (i < pngFiles.Length - 1)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document as a single PDF
        string outputPdf = "Output.pdf";
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Validate that the PDF was created
        if (!File.Exists(outputPdf) || new FileInfo(outputPdf).Length == 0)
            throw new InvalidOperationException("PDF conversion failed: output file was not created.");

        Console.WriteLine($"Successfully created PDF '{outputPdf}' with {pngFiles.Length} pages.");
    }

    // Helper method to generate a PNG image with a solid background and centered text
    private static void CreateSamplePng(string filePath, Aspose.Drawing.Color backgroundColor, string text)
    {
        const int width = 400;
        const int height = 300;

        using (Bitmap bitmap = new Bitmap(width, height))
        {
            // Fill background
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                using (SolidBrush brush = new SolidBrush(backgroundColor))
                {
                    graphics.FillRectangle(brush, 0, 0, width, height);
                }

                // Draw text
                Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24);
                using (SolidBrush textBrush = new SolidBrush(Aspose.Drawing.Color.White))
                {
                    // Measure text size to center it
                    SizeF textSize = graphics.MeasureString(text, font);
                    float x = (width - textSize.Width) / 2;
                    float y = (height - textSize.Height) / 2;
                    graphics.DrawString(text, font, textBrush, new PointF(x, y));
                }
            }

            // Save as PNG
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create output folder
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Generate sample TIFF images
        string[] tiffFiles = new string[2];
        for (int i = 0; i < tiffFiles.Length; i++)
        {
            string tiffPath = Path.Combine(outputDir, $"sample{i + 1}.tiff");
            CreateSampleTiff(tiffPath, i);
            tiffFiles[i] = tiffPath;
        }

        // Create a new document and a builder
        Document pdfDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(pdfDoc);

        // Retrieve page dimensions (in points) and cast to float
        float pageWidth = (float)pdfDoc.FirstSection.PageSetup.PageWidth;
        float pageHeight = (float)pdfDoc.FirstSection.PageSetup.PageHeight;

        // Insert each TIFF as a full‑page image
        for (int i = 0; i < tiffFiles.Length; i++)
        {
            Shape imageShape = builder.InsertImage(tiffFiles[i]);

            // Resize shape to fill the page
            imageShape.Width = pageWidth;
            imageShape.Height = pageHeight;

            // Add a page break after each image except the last one
            if (i < tiffFiles.Length - 1)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document as PDF
        string pdfPath = Path.Combine(outputDir, "ImagesToPdf.pdf");
        pdfDoc.Save(pdfPath, SaveFormat.Pdf);

        // Validate output
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Optional cleanup of temporary TIFF files
        //foreach (var file in tiffFiles) File.Delete(file);
    }

    private static void CreateSampleTiff(string filePath, int index)
    {
        // Create a 500x500 bitmap using Aspose.Drawing
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(500, 500))
        {
            // Obtain a graphics object for drawing
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                // Fill background with white
                graphics.Clear(Aspose.Drawing.Color.White);

                // Choose rectangle color based on index
                Aspose.Drawing.Color rectColor = (index % 2 == 0) ? Aspose.Drawing.Color.LightBlue : Aspose.Drawing.Color.LightCoral;
                using (Aspose.Drawing.Brush brush = new Aspose.Drawing.SolidBrush(rectColor))
                {
                    graphics.FillRectangle(brush, 50, 50, 400, 400);
                }

                // Draw index text using Aspose.Drawing.Font
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 48))
                using (Aspose.Drawing.Brush textBrush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black))
                {
                    graphics.DrawString($"Image {index + 1}", font, textBrush, new Aspose.Drawing.PointF(100, 220));
                }
            }

            // Save bitmap as TIFF
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Tiff);
        }
    }
}

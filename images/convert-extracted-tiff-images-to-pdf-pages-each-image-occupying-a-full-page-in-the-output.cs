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
        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create sample TIFF images
        string[] tiffFiles = new string[2];
        for (int i = 0; i < tiffFiles.Length; i++)
        {
            string tiffPath = Path.Combine(outputDir, $"sample{i + 1}.tiff");
            CreateSampleTiff(tiffPath, i);
            tiffFiles[i] = tiffPath;
        }

        // Create a new Word document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each TIFF image on a separate full page
        for (int i = 0; i < tiffFiles.Length; i++)
        {
            Shape shape = builder.InsertImage(tiffFiles[i]);

            // Make the image fill the page
            shape.WrapType = WrapType.None;
            shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            shape.HorizontalAlignment = HorizontalAlignment.Center;
            shape.VerticalAlignment = VerticalAlignment.Center;

            // Set size to page dimensions (including margins)
            shape.Width = doc.FirstSection.PageSetup.PageWidth;
            shape.Height = doc.FirstSection.PageSetup.PageHeight;

            // Add a page break after each image except the last one
            if (i < tiffFiles.Length - 1)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document as PDF
        string pdfPath = Path.Combine(outputDir, "ImagesToPdf.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created
        if (!File.Exists(pdfPath))
            throw new Exception("PDF file was not created.");
    }

    private static void CreateSampleTiff(string filePath, int index)
    {
        // Create a bitmap with deterministic size
        int width = 600;
        int height = 800;
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);

        // Fill background with a different color per image
        Aspose.Drawing.Color background = (index % 2 == 0) ? Aspose.Drawing.Color.LightBlue : Aspose.Drawing.Color.LightGreen;
        graphics.Clear(background);

        // Draw a semi‑transparent rectangle
        using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.FromArgb(128, Aspose.Drawing.Color.Red)))
        {
            graphics.FillRectangle(brush, 100, 100, 400, 600);
        }

        // Save as TIFF
        bitmap.Save(filePath, ImageFormat.Tiff);

        // Clean up drawing objects
        graphics.Dispose();
        bitmap.Dispose();
    }
}

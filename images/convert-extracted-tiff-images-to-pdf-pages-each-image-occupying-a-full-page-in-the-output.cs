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
        // Folder for generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create sample TIFF images.
        string[] tiffFiles = new string[2];
        for (int i = 0; i < tiffFiles.Length; i++)
        {
            string filePath = Path.Combine(outputDir, $"sample{i + 1}.tiff");
            CreateSampleTiff(filePath, $"Page {i + 1}");
            tiffFiles[i] = filePath;
        }

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each TIFF image on a separate full‑page.
        for (int i = 0; i < tiffFiles.Length; i++)
        {
            if (i > 0)
                builder.InsertBreak(BreakType.PageBreak);

            // Insert the image.
            Shape shape = builder.InsertImage(tiffFiles[i]);

            // Resize the shape to fill the page.
            shape.Width = builder.PageSetup.PageWidth;
            shape.Height = builder.PageSetup.PageHeight;
            shape.WrapType = WrapType.None;
        }

        // Save the document as PDF.
        string pdfPath = Path.Combine(outputDir, "Result.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Optional: clean up temporary TIFF files.
        foreach (string file in tiffFiles)
        {
            if (File.Exists(file))
                File.Delete(file);
        }
    }

    // Creates a simple TIFF image with a solid background and centered text.
    private static void CreateSampleTiff(string filePath, string text)
    {
        const int width = 600;
        const int height = 800;

        // Create bitmap.
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            // Draw background.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.LightGray);
                // Draw text.
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 48, FontStyle.Bold))
                {
                    // Measure text size.
                    SizeF textSize = graphics.MeasureString(text, font);
                    float x = (width - textSize.Width) / 2;
                    float y = (height - textSize.Height) / 2;
                    graphics.DrawString(text, font, Aspose.Drawing.Brushes.Black, x, y);
                }
            }

            // Save as TIFF.
            bitmap.Save(filePath, ImageFormat.Tiff);
        }
    }
}

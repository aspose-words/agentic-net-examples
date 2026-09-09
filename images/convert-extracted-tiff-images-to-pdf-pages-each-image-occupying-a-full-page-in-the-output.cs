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
        // Folder for generated artifacts
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create sample TIFF images (deterministic local files)
        // -----------------------------------------------------------------
        string[] tiffFiles = new string[2];
        for (int i = 0; i < tiffFiles.Length; i++)
        {
            string filePath = Path.Combine(artifactsDir, $"sample{i + 1}.tiff");
            CreateSampleTiff(filePath, $"Page {i + 1}");
            tiffFiles[i] = filePath;
        }

        // -----------------------------------------------------------------
        // 2. Build a new document where each TIFF occupies a full page
        // -----------------------------------------------------------------
        Document pdfDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(pdfDoc);

        // Retrieve page dimensions (in points) from the document's first section
        double pageWidth = pdfDoc.FirstSection.PageSetup.PageWidth;
        double pageHeight = pdfDoc.FirstSection.PageSetup.PageHeight;

        for (int i = 0; i < tiffFiles.Length; i++)
        {
            if (i > 0)
                builder.InsertBreak(BreakType.PageBreak); // start a new page for subsequent images

            // Insert the TIFF image
            Shape imageShape = builder.InsertImage(tiffFiles[i]);

            // Ensure the image fills the whole page
            imageShape.WrapType = WrapType.None;
            imageShape.BehindText = false;
            imageShape.Width = pageWidth;
            imageShape.Height = pageHeight;
        }

        // -----------------------------------------------------------------
        // 3. Save the document as PDF
        // -----------------------------------------------------------------
        string pdfPath = Path.Combine(artifactsDir, "ImagesToPdf.pdf");
        pdfDoc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("PDF output was not created successfully.");

        // Cleanup: optional removal of temporary TIFF files
        foreach (string tiff in tiffFiles)
        {
            if (File.Exists(tiff))
                File.Delete(tiff);
        }
    }

    // Helper method to create a simple single‑frame TIFF image with text
    private static void CreateSampleTiff(string filePath, string caption)
    {
        const int width = 600;
        const int height = 800;

        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                // Fill background
                g.Clear(Aspose.Drawing.Color.White);

                // Draw a rectangle border
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }

                // Draw caption text
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 48, Aspose.Drawing.FontStyle.Bold))
                {
                    SizeF textSize = g.MeasureString(caption, font);
                    float x = (width - textSize.Width) / 2;
                    float y = (height - textSize.Height) / 2;
                    g.DrawString(caption, font, Aspose.Drawing.Brushes.Black, x, y);
                }
            }

            // Save as TIFF (single frame)
            bitmap.Save(filePath, ImageFormat.Tiff);
        }
    }
}

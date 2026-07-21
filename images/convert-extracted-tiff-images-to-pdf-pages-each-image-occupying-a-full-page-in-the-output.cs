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
        // Prepare output folders.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create sample TIFF images.
        int imageCount = 3;
        string[] tiffFiles = new string[imageCount];
        for (int i = 0; i < imageCount; i++)
        {
            string filePath = Path.Combine(artifactsDir, $"sample{i + 1}.tif");
            using (Bitmap bitmap = new Bitmap(600, 800))
            {
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    // Fill background with a distinct color.
                    g.Clear(i % 2 == 0 ? Color.LightBlue : Color.LightGreen);
                    // Optionally draw simple text.
                    // (Aspose.Drawing does not provide a direct DrawString overload without a Font,
                    //  so we keep the image simple.)
                }
                bitmap.Save(filePath, ImageFormat.Tiff);
            }
            tiffFiles[i] = filePath;
        }

        // Verify that images were created.
        if (tiffFiles.Length == 0 || !File.Exists(tiffFiles[0]))
            throw new InvalidOperationException("No TIFF images were created.");

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Page dimensions (excluding margins) for full‑page image placement.
        double pageWidth = doc.FirstSection.PageSetup.PageWidth - doc.FirstSection.PageSetup.LeftMargin - doc.FirstSection.PageSetup.RightMargin;
        double pageHeight = doc.FirstSection.PageSetup.PageHeight - doc.FirstSection.PageSetup.TopMargin - doc.FirstSection.PageSetup.BottomMargin;

        // Insert each TIFF image on its own page.
        for (int i = 0; i < tiffFiles.Length; i++)
        {
            // Insert the image and obtain the Shape object.
            Shape imgShape = builder.InsertImage(tiffFiles[i]);

            // Resize the shape to fill the printable area of the page.
            imgShape.Width = pageWidth;
            imgShape.Height = pageHeight;

            // Position the image relative to the page and disable text wrapping.
            imgShape.WrapType = WrapType.None;
            imgShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            imgShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            imgShape.HorizontalAlignment = HorizontalAlignment.Center;
            imgShape.VerticalAlignment = VerticalAlignment.Center;

            // Add a page break after each image except the last one.
            if (i < tiffFiles.Length - 1)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document as PDF.
        string pdfPath = Path.Combine(artifactsDir, "ImagesToPdf.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // The program finishes here without waiting for user input.
    }
}

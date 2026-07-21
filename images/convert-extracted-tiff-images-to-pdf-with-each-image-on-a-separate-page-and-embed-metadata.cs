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
        // Define file names for the sample TIFF images and the output PDF.
        string[] tiffFiles = { "sample1.tiff", "sample2.tiff" };
        string pdfFile = "ConvertedImages.pdf";

        // Create deterministic sample TIFF images.
        CreateSampleTiff(tiffFiles[0], 200, 200, Aspose.Drawing.Color.LightBlue, "Image 1");
        CreateSampleTiff(tiffFiles[1], 300, 150, Aspose.Drawing.Color.LightGreen, "Image 2");

        // Create a new Word document and a builder to insert content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each TIFF on its own page.
        for (int i = 0; i < tiffFiles.Length; i++)
        {
            builder.InsertImage(tiffFiles[i]);

            // Add a page break after each image except the last one.
            if (i < tiffFiles.Length - 1)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Embed built‑in metadata.
        doc.BuiltInDocumentProperties.Title = "Converted PDF from TIFF images";
        doc.BuiltInDocumentProperties.Author = "Aspose.Words Example";
        doc.BuiltInDocumentProperties.Subject = "TIFF to PDF conversion";

        // Add a custom document property.
        doc.CustomDocumentProperties.Add("Source", "Generated sample TIFF images");

        // Save the document as PDF.
        doc.Save(pdfFile, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(pdfFile))
            throw new InvalidOperationException($"Failed to create PDF file: {pdfFile}");

        // Clean up temporary TIFF files.
        foreach (string file in tiffFiles)
        {
            if (File.Exists(file))
                File.Delete(file);
        }
    }

    // Helper method to create a deterministic TIFF image.
    private static void CreateSampleTiff(string fileName, int width, int height, Aspose.Drawing.Color backColor, string text)
    {
        // Create a bitmap with the requested dimensions.
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            // Obtain a graphics object for drawing.
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                // Fill the background.
                graphics.Clear(backColor);

                // Draw simple text.
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20))
                {
                    graphics.DrawString(text, font, Aspose.Drawing.Brushes.Black, new Aspose.Drawing.PointF(10, height / 2 - 10));
                }
            }

            // Save the bitmap as a TIFF file.
            bitmap.Save(fileName, ImageFormat.Tiff);
        }
    }
}

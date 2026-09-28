using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare sample TIFF images.
        string[] tiffFiles = { "sample1.tif", "sample2.tif" };
        int width = 400;
        int height = 300;

        for (int i = 0; i < tiffFiles.Length; i++)
        {
            // Create a bitmap and draw deterministic content.
            using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
            {
                using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
                {
                    g.Clear(Aspose.Drawing.Color.White);
                    // Draw a simple rectangle with text indicating the image number.
                    using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 3))
                    {
                        g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                    }

                    using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black))
                    {
                        using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
                        {
                            g.DrawString($"Image {i + 1}", font, brush, new Aspose.Drawing.PointF(50, height / 2 - 20));
                        }
                    }
                }

                // Save as TIFF.
                bitmap.Save(tiffFiles[i], ImageFormat.Tiff);
            }
        }

        // Verify that TIFF files were created.
        foreach (string file in tiffFiles)
        {
            if (!File.Exists(file))
                throw new Exception($"Failed to create TIFF image: {file}");
        }

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each TIFF image on a separate page.
        for (int i = 0; i < tiffFiles.Length; i++)
        {
            if (i > 0)
                builder.InsertBreak(BreakType.PageBreak);

            // Insert the image.
            builder.InsertImage(tiffFiles[i]);
        }

        // Embed metadata.
        doc.BuiltInDocumentProperties.Title = "Converted TIFF Images to PDF";
        doc.BuiltInDocumentProperties.Author = "Aspose.Words Example";
        doc.CustomDocumentProperties.Add("SourceFormat", "TIFF");
        doc.CustomDocumentProperties.Add("ImageCount", tiffFiles.Length);

        // Save the document as PDF.
        string outputPdf = "output.pdf";
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPdf))
            throw new Exception("PDF output was not created.");

        // Clean up sample TIFF files (optional).
        foreach (string file in tiffFiles)
        {
            try { File.Delete(file); } catch { /* ignore */ }
        }
    }
}

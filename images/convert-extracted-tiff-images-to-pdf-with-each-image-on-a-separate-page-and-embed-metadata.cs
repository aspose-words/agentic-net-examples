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
        // Create deterministic sample TIFF images.
        CreateSampleTiff("sample1.tif", "Page 1");
        CreateSampleTiff("sample2.tif", "Page 2");

        // Verify that the TIFF files were created.
        if (!File.Exists("sample1.tif") || !File.Exists("sample2.tif"))
            throw new FileNotFoundException("Sample TIFF images were not created.");

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first TIFF image.
        builder.InsertImage("sample1.tif");
        // Insert a page break to start a new page.
        builder.InsertBreak(BreakType.PageBreak);
        // Insert the second TIFF image.
        builder.InsertImage("sample2.tif");

        // Embed metadata into the PDF.
        doc.BuiltInDocumentProperties.Title = "Combined PDF from TIFFs";
        doc.BuiltInDocumentProperties.Author = "Aspose.Words Example";
        doc.CustomDocumentProperties.Add("Source", "Generated sample TIFF images");

        // Save the document as PDF.
        string pdfPath = "output.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new Exception("PDF file was not created.");

        Console.WriteLine($"PDF successfully created at '{Path.GetFullPath(pdfPath)}'.");
    }

    private static void CreateSampleTiff(string fileName, string text)
    {
        // Define image size.
        int width = 400;
        int height = 300;

        // Create a bitmap and draw deterministic content.
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.White);

            // Use Aspose.Drawing.Font explicitly to avoid ambiguity.
            using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
            using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black))
            {
                graphics.DrawString(text, font, brush, new Aspose.Drawing.PointF(10, height / 2 - 20));
            }

            // Save as a single‑frame TIFF image.
            bitmap.Save(fileName, Aspose.Drawing.Imaging.ImageFormat.Tiff);
        }
    }
}

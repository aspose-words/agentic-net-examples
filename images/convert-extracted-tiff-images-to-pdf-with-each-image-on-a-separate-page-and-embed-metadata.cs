using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ConvertTiffToPdf
{
    public static void Main()
    {
        // Define file names
        string tiff1Path = "sample1.tif";
        string tiff2Path = "sample2.tif";
        string pdfPath = "ConvertedImages.pdf";

        // -------------------------------------------------
        // 1. Create sample TIFF images (deterministic local files)
        // -------------------------------------------------
        CreateSampleTiff(tiff1Path, Aspose.Drawing.Color.LightCoral, "Image 1");
        CreateSampleTiff(tiff2Path, Aspose.Drawing.Color.LightGreen, "Image 2");

        // Verify that the TIFF files were created
        if (!File.Exists(tiff1Path) || !File.Exists(tiff2Path))
            throw new FileNotFoundException("Failed to create sample TIFF images.");

        // -------------------------------------------------
        // 2. Create a new Word document and insert each TIFF on a separate page
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert first image
        builder.InsertImage(tiff1Path);
        // Add a page break before the next image (if any)
        builder.InsertBreak(BreakType.PageBreak);

        // Insert second image
        builder.InsertImage(tiff2Path);

        // -------------------------------------------------
        // 3. Embed metadata into the document
        // -------------------------------------------------
        doc.BuiltInDocumentProperties.Title = "Converted TIFF Images to PDF";
        doc.BuiltInDocumentProperties.Author = "Aspose.Words Example";
        doc.BuiltInDocumentProperties.Keywords = "TIFF, PDF, conversion, Aspose.Words";

        // -------------------------------------------------
        // 4. Save the document as PDF
        // -------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Ensure metadata is written to the PDF
            ExportDocumentStructure = true
        };
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF was created
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF conversion failed.");

        // Cleanup temporary TIFF files (optional)
        File.Delete(tiff1Path);
        File.Delete(tiff2Path);
    }

    private static void CreateSampleTiff(string filePath, Aspose.Drawing.Color backgroundColor, string text)
    {
        // Create a bitmap and draw deterministic content
        int width = 400;
        int height = 300;
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(backgroundColor);

        // Simple text drawing (optional, uses default font)
        // Use fully qualified Aspose.Drawing.Font to avoid ambiguity
        using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
        using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black))
        {
            graphics.DrawString(text, font, brush, new Aspose.Drawing.PointF(10, height / 2 - 12));
        }

        // Save as TIFF
        bitmap.Save(filePath, ImageFormat.Tiff);

        // Release resources
        graphics.Dispose();
        bitmap.Dispose();
    }
}

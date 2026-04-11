using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for temporary files
        string imagePath = Path.Combine(outputDir, "sample.png");
        string docxPath = Path.Combine(outputDir, "sample.docx");
        string pdfPath = Path.Combine(outputDir, "sample.pdf");

        // Create a simple PNG image using Aspose.Drawing
        int imgWidth = 200;
        int imgHeight = 200;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
                using (Pen pen = new Pen(Color.Red, 5))
                {
                    graphics.DrawEllipse(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Verify that the image was created
        if (!File.Exists(imagePath) || new FileInfo(imagePath).Length == 0)
            throw new InvalidOperationException("Failed to create the sample image.");

        // Create a DOCX document and insert the image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document with an image:");
        builder.InsertImage(imagePath);
        doc.Save(docxPath, SaveFormat.Docx);

        // Load the DOCX document (bootstrap rule)
        Document loadedDoc = new Document(docxPath);

        // Configure PDF save options for high image compression
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ImageCompression = PdfImageCompression.Jpeg, // Compress all images as JPEG
            JpegQuality = 10 // Low quality = high compression
        };

        // Save the document as PDF
        loadedDoc.Save(pdfPath, pdfOptions);

        // Validate that the PDF was created and is not empty
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("PDF conversion failed or produced an empty file.");
    }
}

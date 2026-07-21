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
        // Prepare folders
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample image using Aspose.Drawing
        byte[] imageBytes;
        using (var bitmap = new Bitmap(100, 100))
        {
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                using (var pen = new Pen(Color.Blue, 3))
                {
                    graphics.DrawEllipse(pen, new Rectangle(10, 10, 80, 80));
                }
            }

            using (var ms = new MemoryStream())
            {
                bitmap.Save(ms, ImageFormat.Png);
                imageBytes = ms.ToArray();
            }
        }

        // Create a Word document and insert the image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document with an embedded image:");
        using (var imgStream = new MemoryStream(imageBytes))
        {
            builder.InsertImage(imgStream);
        }

        // Save the document as PDF
        string pdfPath = Path.Combine(outputDir, "sample.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Load the PDF document
        Document pdfDoc = new Document(pdfPath);

        // Save PDF as HTML with images embedded as Base64
        string htmlPath = Path.Combine(outputDir, "sample.html");
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportEmbeddedImages = true,
            PrettyFormat = true
        };
        pdfDoc.Save(htmlPath, htmlOptions);
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML file was not created.");

        // Verify that the HTML contains Base64 image data
        string htmlContent = File.ReadAllText(htmlPath);
        if (!htmlContent.Contains("data:image"))
            throw new InvalidOperationException("Images were not embedded as Base64 in the HTML.");

        // Example completed successfully
        Console.WriteLine("PDF converted to HTML with embedded Base64 images.");
    }
}

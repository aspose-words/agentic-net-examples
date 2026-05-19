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
        // Prepare output directories
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple PNG image using Aspose.Drawing
        string imagePath = Path.Combine(artifactsDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Create a Word document, insert the image, and save it as PDF
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF with an embedded image.");
        builder.InsertImage(imagePath);
        string pdfPath = Path.Combine(artifactsDir, "sample.pdf");
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document
        Document pdfDoc = new Document(pdfPath);

        // Configure HtmlFixedSaveOptions to embed images as Base64 data URIs
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportEmbeddedImages = true,
            PrettyFormat = true
        };

        // Save the PDF as HTML with embedded images
        string htmlPath = Path.Combine(artifactsDir, "output.html");
        pdfDoc.Save(htmlPath, htmlOptions);

        // Validate that the HTML file was created
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("Expected output HTML was not created.");

        // Validate that the HTML contains Base64-encoded image data
        string htmlContent = File.ReadAllText(htmlPath);
        if (!htmlContent.Contains("data:image"))
            throw new InvalidOperationException("HTML does not contain Base64-encoded image data.");

        // Clean up temporary files (optional)
        // File.Delete(imagePath);
        // File.Delete(pdfPath);
    }
}

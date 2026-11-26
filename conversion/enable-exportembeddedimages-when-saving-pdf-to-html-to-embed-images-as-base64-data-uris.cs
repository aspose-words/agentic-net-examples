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
        // Create a sample document with an image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document with an image.");

        // Create a simple bitmap using Aspose.Drawing.
        string imagePath = "sample.png";
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Insert the image into the document.
        builder.InsertImage(imagePath);

        // Save the document as PDF (input.pdf).
        string pdfPath = "input.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Configure HTML save options to embed images as Base64 data URIs.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportImagesAsBase64 = true
        };

        // Save the PDF as HTML.
        string htmlPath = "output.html";
        pdfDoc.Save(htmlPath, htmlOptions);

        // Verify that the HTML file exists and contains embedded image data.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML file was not created.");

        string htmlContent = File.ReadAllText(htmlPath);
        if (!htmlContent.Contains("data:image"))
            throw new InvalidOperationException("Embedded image data URI not found in HTML.");

        // Clean up temporary image file.
        File.Delete(imagePath);
    }
}

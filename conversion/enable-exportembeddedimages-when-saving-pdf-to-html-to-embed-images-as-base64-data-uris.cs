using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define output directories and file names.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        string pdfPath = Path.Combine(outputDir, "Sample.pdf");
        string htmlPath = Path.Combine(outputDir, "Sample.html");

        // -----------------------------------------------------------------
        // 1. Create a simple Word document with text and an image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document will be saved as PDF and then converted to HTML with embedded images.");

        // Create a red square image using Aspose.Drawing (no System.Drawing usage).
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Red);
            }

            using (MemoryStream imageStream = new MemoryStream())
            {
                // Save the bitmap to a PNG stream.
                bitmap.Save(imageStream, ImageFormat.Png);
                imageStream.Position = 0; // Reset for reading.

                // Insert the image into the document.
                builder.InsertImage(imageStream);
            }
        }

        // -----------------------------------------------------------------
        // 2. Save the document as PDF (the source for conversion).
        // -----------------------------------------------------------------
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify PDF creation.
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("Failed to create the PDF file.");

        // -----------------------------------------------------------------
        // 3. Load the PDF and convert it to HTML with images embedded as Base64.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportEmbeddedImages = true, // Embed images as Base64 data URIs.
            PrettyFormat = true
        };

        pdfDoc.Save(htmlPath, htmlOptions);

        // -----------------------------------------------------------------
        // 4. Validate that the HTML file exists and contains Base64 image data.
        // -----------------------------------------------------------------
        if (!File.Exists(htmlPath) || new FileInfo(htmlPath).Length == 0)
            throw new InvalidOperationException("HTML conversion failed; output file is missing or empty.");

        string htmlContent = File.ReadAllText(htmlPath, Encoding.UTF8);
        if (!htmlContent.Contains("data:image"))
            throw new InvalidOperationException("HTML does not contain embedded Base64 images.");

        // Example completed successfully.
        Console.WriteLine("PDF successfully converted to HTML with embedded Base64 images.");
    }
}

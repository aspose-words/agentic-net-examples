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
        // Define paths for temporary files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        string imagePath = Path.Combine(artifactsDir, "sample.png");
        string pdfPath = Path.Combine(artifactsDir, "sample.pdf");
        string htmlPath = Path.Combine(artifactsDir, "sample.html");

        // -----------------------------------------------------------------
        // Create a simple PNG image using Aspose.Drawing (no System.Drawing).
        // -----------------------------------------------------------------
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill the bitmap with a solid color.
                graphics.Clear(Color.Blue);
            }

            // Save the bitmap to a file.
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // --------------------------------------------------------------
        // Create a Word document, insert the image, and save it as PDF.
        // --------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample PDF with an embedded image:");
        builder.InsertImage(imagePath);
        doc.Save(pdfPath, SaveFormat.Pdf);

        // --------------------------------------------------------------
        // Load the PDF and save it as HTML with images embedded as Base64.
        // --------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportEmbeddedImages = true, // Enable Base64 embedding.
            PrettyFormat = true
        };
        pdfDoc.Save(htmlPath, htmlOptions);

        // ------------------------------
        // Validate that the HTML contains Base64 image data.
        // ------------------------------
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("The HTML output file was not created.");

        string htmlContent = File.ReadAllText(htmlPath);
        if (!htmlContent.Contains("data:image"))
            throw new InvalidOperationException("Images were not embedded as Base64 data URIs.");

        // Optional: clean up temporary files (comment out if inspection is needed).
        // File.Delete(imagePath);
        // File.Delete(pdfPath);
        // File.Delete(htmlPath);
        // Directory.Delete(artifactsDir, true);
    }
}

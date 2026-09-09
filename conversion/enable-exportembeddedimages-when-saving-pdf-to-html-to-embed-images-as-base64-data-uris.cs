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
        // Define folders for generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a simple image using Aspose.Drawing and save it to disk.
        // -----------------------------------------------------------------
        string imagePath = Path.Combine(artifactsDir, "sample.png");
        CreateSamplePng(imagePath);

        // -----------------------------------------------------------------
        // 2. Build a Word document and insert the created image.
        // -----------------------------------------------------------------
        Document wordDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(wordDoc);
        builder.Writeln("Document with an embedded image:");
        builder.InsertImage(imagePath);
        string pdfPath = Path.Combine(artifactsDir, "sample.pdf");
        wordDoc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 3. Load the PDF document.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // -----------------------------------------------------------------
        // 4. Save the PDF as fixed HTML with images embedded as Base64.
        // -----------------------------------------------------------------
        string htmlPath = Path.Combine(artifactsDir, "sample.html");
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportEmbeddedImages = true,   // Embed images as Base64 data URIs.
            PrettyFormat = true
        };
        pdfDoc.Save(htmlPath, htmlOptions);

        // -----------------------------------------------------------------
        // 5. Validate that the HTML file was created and contains Base64 image data.
        // -----------------------------------------------------------------
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML output file was not created.");

        string htmlContent = File.ReadAllText(htmlPath);
        if (!htmlContent.Contains("data:image"))
            throw new InvalidOperationException("ExportEmbeddedImages did not embed images as Base64.");

        // Example completed successfully.
    }

    // Creates a 100x100 PNG with a simple drawing using Aspose.Drawing.
    private static void CreateSamplePng(string filePath)
    {
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightGray);
                using (SolidBrush brush = new SolidBrush(Color.Blue))
                {
                    graphics.FillEllipse(brush, new Rectangle(10, 10, 80, 80));
                }
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}

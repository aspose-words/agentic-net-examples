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
        // Prepare file paths.
        string pdfPath = "sample.pdf";
        string htmlPath = "output.html";

        // -----------------------------------------------------------------
        // 1. Create a simple Word document with an image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document with an embedded image.");

        // Create a 100x100 red square using Aspose.Drawing.
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Red);
            }

            // Save the bitmap to a memory stream in PNG format.
            using (MemoryStream imageStream = new MemoryStream())
            {
                bitmap.Save(imageStream, ImageFormat.Png);
                imageStream.Position = 0;

                // Insert the image into the document.
                builder.InsertImage(imageStream);
            }
        }

        // -----------------------------------------------------------------
        // 2. Save the document as PDF.
        // -----------------------------------------------------------------
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify PDF creation.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // -----------------------------------------------------------------
        // 3. Load the PDF and save it as HTML with embedded images.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportEmbeddedImages = true, // Embed images as Base64 data URIs.
            PrettyFormat = true
        };

        pdfDoc.Save(htmlPath, htmlOptions);

        // Verify HTML creation.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML file was not created.");

        // -----------------------------------------------------------------
        // 4. Validate that the HTML contains Base64 image data.
        // -----------------------------------------------------------------
        string htmlContent = File.ReadAllText(htmlPath);
        if (!htmlContent.Contains("data:image"))
            throw new InvalidOperationException("ExportEmbeddedImages did not embed images as Base64.");

        // If execution reaches this point, the process succeeded.
        Console.WriteLine("PDF successfully converted to HTML with embedded Base64 images.");
    }
}

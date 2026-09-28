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
        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document with an embedded image.");

        // Create a simple bitmap image using Aspose.Drawing.
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }

            using (MemoryStream imageStream = new MemoryStream())
            {
                // Save bitmap to stream as PNG.
                bitmap.Save(imageStream, ImageFormat.Png);
                imageStream.Position = 0;

                // Insert the image into the document.
                builder.InsertImage(imageStream);
            }
        }

        // Configure HTML save options to embed images as Base64 data URIs.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportImagesAsBase64 = true;

        string outputPath = "output.html";
        doc.Save(outputPath, saveOptions);

        // Validate that the HTML file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output HTML file was not created.");
        }

        // Verify that the HTML contains a data URI for the embedded image.
        string htmlContent = File.ReadAllText(outputPath);
        if (!htmlContent.Contains("data:image"))
        {
            throw new InvalidOperationException("Embedded images were not exported as Base64 data URIs.");
        }
    }
}

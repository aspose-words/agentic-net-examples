using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    // Model for simulated REST API response
    private class ImageResponse
    {
        public string ImageData { get; set; }
    }

    public static void Main()
    {
        // Paths for files used in the example
        const string templatePath = "template.docx";
        const string sampleImagePath = "sample.png";
        const string outputPdfPath = "output.pdf";

        // -------------------------------------------------
        // 1. Create a sample image (simulating an image from a REST API)
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                // Draw a simple red ellipse
                using (Pen pen = new Pen(Color.Red, 5))
                {
                    graphics.DrawEllipse(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }
            // Save the image locally
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // 2. Simulate a REST API that returns the image as Base64 JSON
        // -------------------------------------------------
        byte[] imageBytes = File.ReadAllBytes(sampleImagePath);
        string base64Image = Convert.ToBase64String(imageBytes);
        string jsonResponse = JsonConvert.SerializeObject(new ImageResponse { ImageData = base64Image });

        // -------------------------------------------------
        // 3. Parse the simulated REST API response
        // -------------------------------------------------
        ImageResponse response = JsonConvert.DeserializeObject<ImageResponse>(jsonResponse);
        if (response == null || string.IsNullOrEmpty(response.ImageData))
            throw new Exception("Failed to retrieve image data from simulated API.");

        byte[] apiImageBytes = Convert.FromBase64String(response.ImageData);

        // -------------------------------------------------
        // 4. Create a DOCX template (if it does not already exist)
        // -------------------------------------------------
        if (!File.Exists(templatePath))
        {
            Document templateDoc = new Document();
            DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);
            tmplBuilder.Writeln("This is a template document.");
            tmplBuilder.Writeln("Below is the image fetched from the REST API:");
            templateDoc.Save(templatePath);
        }

        // -------------------------------------------------
        // 5. Load the DOCX template and insert the image
        // -------------------------------------------------
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move builder to the end of the document
        builder.MoveToDocumentEnd();

        // Insert a paragraph break before the image
        builder.Writeln();

        // Insert image from the byte array using a MemoryStream
        using (MemoryStream ms = new MemoryStream(apiImageBytes))
        {
            ms.Position = 0; // Ensure the stream is at the beginning
            builder.InsertImage(ms);
        }

        // -------------------------------------------------
        // 6. Save the resulting document as PDF
        // -------------------------------------------------
        doc.Save(outputPdfPath, SaveFormat.Pdf);

        // -------------------------------------------------
        // 7. Validate that the PDF was created
        // -------------------------------------------------
        if (!File.Exists(outputPdfPath))
            throw new Exception("PDF output was not created.");

        // Optional: clean up temporary files (comment out if you want to inspect them)
        // File.Delete(sampleImagePath);
        // File.Delete(templatePath);
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    // Helper class for JSON deserialization
    private class ImageResponse
    {
        public string imageData { get; set; }
    }

    public static void Main()
    {
        // Paths for files used in the example
        const string templatePath = "template.docx";
        const string outputPdfPath = "output.pdf";

        // -------------------------------------------------
        // 1. Create a simple DOCX template if it does not exist
        // -------------------------------------------------
        if (!File.Exists(templatePath))
        {
            Document templateDoc = new Document();
            DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
            templateBuilder.Writeln("Template Document");
            templateBuilder.Writeln("The image will be inserted below:");
            templateDoc.Save(templatePath);
        }

        // -------------------------------------------------
        // 2. Simulate a REST API that returns an image as Base64 JSON
        // -------------------------------------------------
        // Create a deterministic sample image using Aspose.Drawing
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.LightBlue);
        // Draw a simple rectangle
        using (var pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.DarkBlue, 5))
        {
            graphics.DrawRectangle(pen, 20, 20, 160, 160);
        }

        // Save the image to a memory stream in PNG format
        byte[] pngBytes;
        using (MemoryStream imgStream = new MemoryStream())
        {
            bitmap.Save(imgStream, ImageFormat.Png);
            imgStream.Position = 0;
            pngBytes = imgStream.ToArray();
        }

        // Clean up drawing resources
        graphics.Dispose();
        bitmap.Dispose();

        // Encode the image bytes to Base64 and wrap in JSON
        string base64Image = Convert.ToBase64String(pngBytes);
        string jsonResponse = $"{{\"imageData\":\"{base64Image}\"}}";

        // -------------------------------------------------
        // 3. Parse the JSON response to obtain the image bytes
        // -------------------------------------------------
        ImageResponse response = JsonConvert.DeserializeObject<ImageResponse>(jsonResponse);
        if (response == null || string.IsNullOrEmpty(response.imageData))
            throw new InvalidOperationException("Failed to retrieve image data from simulated API.");

        byte[] imageBytes = Convert.FromBase64String(response.imageData);

        // -------------------------------------------------
        // 4. Load the DOCX template, insert the image, and save as PDF
        // -------------------------------------------------
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move to the end of the document to insert the image
        builder.MoveToDocumentEnd();
        builder.InsertParagraph(); // Ensure a new paragraph for the image
        builder.InsertImage(imageBytes);

        // Save the resulting document as PDF
        doc.Save(outputPdfPath, SaveFormat.Pdf);

        // -------------------------------------------------
        // 5. Validate that the PDF was created
        // -------------------------------------------------
        if (!File.Exists(outputPdfPath))
            throw new FileNotFoundException("The PDF output file was not created.", outputPdfPath);

        // Example completed successfully
        Console.WriteLine("Document processed and saved as PDF successfully.");
    }
}

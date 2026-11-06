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
    // Model for the simulated REST API response
    private class ImageResponse
    {
        public string ImageData { get; set; }
    }

    public static void Main()
    {
        // 1. Create a deterministic sample image that will act as the image returned by the REST API
        const string apiImagePath = "apiImage.png";
        CreateSampleImage(apiImagePath);

        // 2. Encode the image to Base64 and build a JSON string that mimics a REST API response
        string base64Image = Convert.ToBase64String(File.ReadAllBytes(apiImagePath));
        string jsonResponse = JsonConvert.SerializeObject(new ImageResponse { ImageData = base64Image });

        // 3. Simulate calling a REST API and parsing the JSON response
        ImageResponse response = JsonConvert.DeserializeObject<ImageResponse>(jsonResponse);
        if (response == null || string.IsNullOrEmpty(response.ImageData))
            throw new InvalidOperationException("Failed to retrieve image data from simulated API.");

        // 4. Decode the Base64 image data and save it to a file that will be inserted into the document
        byte[] imageBytes = Convert.FromBase64String(response.ImageData);
        const string downloadedImagePath = "downloadedFromApi.png";
        File.WriteAllBytes(downloadedImagePath, imageBytes);

        // 5. Create a simple DOCX template file
        const string templatePath = "template.docx";
        CreateTemplateDocument(templatePath);

        // 6. Load the DOCX template
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // 7. Insert the image obtained from the simulated REST API into the document
        builder.Writeln(); // Ensure we are on a new paragraph
        builder.InsertImage(downloadedImagePath);

        // 8. Save the final document as PDF
        const string outputPdfPath = "output.pdf";
        doc.Save(outputPdfPath, SaveFormat.Pdf);

        // 9. Validate that the PDF was created
        if (!File.Exists(outputPdfPath))
            throw new FileNotFoundException("The PDF output file was not created.", outputPdfPath);
    }

    // Helper method to create a deterministic sample image using Aspose.Drawing
    private static void CreateSampleImage(string filePath)
    {
        const int width = 200;
        const int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                // Draw a simple rectangle
                using (Pen pen = new Pen(Color.Blue, 5))
                {
                    graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }

        // Ensure the image file exists
        if (!File.Exists(filePath))
            throw new FileNotFoundException("Failed to create sample image.", filePath);
    }

    // Helper method to create a minimal DOCX template
    private static void CreateTemplateDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a DOCX template. The image will be inserted below:");
        doc.Save(filePath, SaveFormat.Docx);

        // Ensure the template file exists
        if (!File.Exists(filePath))
            throw new FileNotFoundException("Failed to create DOCX template.", filePath);
    }
}

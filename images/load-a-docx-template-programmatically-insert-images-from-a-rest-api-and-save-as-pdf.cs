using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // File names used in the example
        const string templatePath = "template.docx";
        const string imagePath = "apiImage.png";
        const string outputPdfPath = "output.pdf";

        // -------------------------------------------------
        // Step 1: Create a sample image (simulating a REST API response)
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 100;

        // Create a deterministic bitmap using Aspose.Drawing
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.LightBlue);
        // Draw a simple rectangle to make the image visible
        graphics.FillRectangle(
            new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.DarkBlue),
            10, 10, imgWidth - 20, imgHeight - 20);

        // Save the generated image to a local file
        bitmap.Save(imagePath);

        // Release drawing resources
        graphics.Dispose();
        bitmap.Dispose();

        // -------------------------------------------------
        // Step 2: Create a simple DOCX template programmatically
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("This is a sample template document.");
        templateBuilder.Writeln("Below is the image fetched from the API:");
        // Save the template so it can be loaded later
        templateDoc.Save(templatePath);
        // No explicit Dispose needed for Document or DocumentBuilder

        // -------------------------------------------------
        // Step 3: Load the DOCX template
        // -------------------------------------------------
        Document doc = new Document(templatePath);
        DocumentBuilder docBuilder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // Step 4: Insert the image into the document
        // -------------------------------------------------
        if (!File.Exists(imagePath))
            throw new FileNotFoundException("Image file not found.", imagePath);

        // Insert the image at the current cursor position
        docBuilder.InsertImage(imagePath);

        // -------------------------------------------------
        // Step 5: Save the resulting document as PDF
        // -------------------------------------------------
        doc.Save(outputPdfPath, SaveFormat.Pdf);

        // -------------------------------------------------
        // Validation: Ensure the PDF was created
        // -------------------------------------------------
        if (!File.Exists(outputPdfPath))
            throw new Exception("Failed to create the output PDF file.");

        // Indicate successful completion
        Console.WriteLine("Document processed and saved as PDF successfully.");
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Define file paths
        string templatePath = Path.Combine(artifactsDir, "template.docx");
        string imagePath = Path.Combine(artifactsDir, "apiImage.png");
        string outputPdfPath = Path.Combine(artifactsDir, "result.pdf");

        // 1. Create a sample image (simulating a REST API response)
        CreateSampleImage(imagePath);

        // 2. Create a simple DOCX template
        CreateTemplateDocument(templatePath);

        // 3. Load the template document
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // 4. Insert the image into the document
        builder.InsertParagraph(); // ensure a new paragraph before the image
        builder.InsertImage(imagePath);

        // 5. Save the document as PDF
        doc.Save(outputPdfPath, SaveFormat.Pdf);

        // 6. Validate that the PDF was created
        if (!File.Exists(outputPdfPath))
        {
            throw new InvalidOperationException("PDF output was not created.");
        }
    }

    private static void CreateSampleImage(string filePath)
    {
        // Create a 200x200 bitmap
        Bitmap bitmap = new Bitmap(200, 200);
        Graphics graphics = Graphics.FromImage(bitmap);

        // Fill background with white
        graphics.Clear(Color.White);

        // Draw a simple red rectangle
        using (Pen pen = new Pen(Color.Red, 5))
        {
            graphics.DrawRectangle(pen, 20, 20, 160, 160);
        }

        // Save the bitmap to a PNG file
        bitmap.Save(filePath);

        // Clean up resources
        graphics.Dispose();
        bitmap.Dispose();
    }

    private static void CreateTemplateDocument(string filePath)
    {
        // Create a blank document and add a placeholder paragraph
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("This is a template document.");

        // Save the template
        template.Save(filePath);
    }
}

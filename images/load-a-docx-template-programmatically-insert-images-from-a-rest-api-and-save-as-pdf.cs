using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a simple DOCX template that will be loaded later.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(artifactsDir, "Template.docx");
        CreateTemplateDocument(templatePath);

        // -----------------------------------------------------------------
        // 2. Simulate images retrieved from a REST API by creating local files.
        // -----------------------------------------------------------------
        string imagePath1 = Path.Combine(artifactsDir, "ApiImage1.png");
        string imagePath2 = Path.Combine(artifactsDir, "ApiImage2.png");
        CreateSampleImage(imagePath1, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(imagePath2, Aspose.Drawing.Color.LightGreen);

        // -----------------------------------------------------------------
        // 3. Load the template, insert the images, and save as PDF.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert first image.
        builder.InsertImage(imagePath1);
        builder.Writeln(); // Add a line break between images.

        // Insert second image.
        builder.InsertImage(imagePath2);
        builder.Writeln();

        // Save the resulting document as PDF.
        string pdfPath = Path.Combine(artifactsDir, "Result.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 4. Validate that the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // (Optional) Output paths for verification when running locally.
        Console.WriteLine("Template created at: " + templatePath);
        Console.WriteLine("Image 1 created at: " + imagePath1);
        Console.WriteLine("Image 2 created at: " + imagePath2);
        Console.WriteLine("PDF saved at: " + pdfPath);
    }

    // Creates a minimal DOCX file containing a single line of text.
    private static void CreateTemplateDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a template document. Images will be inserted below:");
        doc.Save(filePath);
    }

    // Generates a simple PNG image with a solid background color.
    private static void CreateSampleImage(string filePath, Aspose.Drawing.Color backgroundColor)
    {
        const int width = 200;
        const int height = 150;

        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(backgroundColor);
            }

            // Ensure the bitmap is saved before disposing.
            bitmap.Save(filePath);
        }
    }
}

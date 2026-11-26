using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample image file.
        string imagePath = "input.png";
        CreateSampleImage(imagePath);
        if (!File.Exists(imagePath))
            throw new Exception("Sample image was not created.");

        // Create a deterministic DOCX template.
        string templatePath = "template.docx";
        CreateTemplateDocument(templatePath);
        if (!File.Exists(templatePath))
            throw new Exception("Template document was not created.");

        // Load the template document.
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image into the document.
        Shape insertedShape = builder.InsertImage(imagePath);
        if (!insertedShape.HasImage)
            throw new Exception("Image insertion failed.");

        // Save the resulting document.
        string outputPath = "output.docx";
        doc.Save(outputPath);
        if (!File.Exists(outputPath))
            throw new Exception("Output document was not created.");
    }

    private static void CreateSampleImage(string path)
    {
        int width = 200;
        int height = 100;

        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                graphics.Clear(Color.White);

                // Draw a simple blue rectangle.
                using (Pen pen = new Pen(Color.Blue, 3))
                {
                    graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }

            // Save the bitmap to a PNG file.
            bitmap.Save(path);
        }
    }

    private static void CreateTemplateDocument(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a template document.");
        doc.Save(path);
    }
}

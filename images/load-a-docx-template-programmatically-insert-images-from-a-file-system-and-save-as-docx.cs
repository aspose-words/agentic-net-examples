using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define deterministic file names.
        const string imagePath = "input.png";
        const string templatePath = "template.docx";
        const string outputPath = "output.docx";

        // -------------------------------------------------
        // Create a sample image using Aspose.Drawing.
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                g.Clear(Color.White);
                // Optionally draw a simple rectangle.
                g.DrawRectangle(Pens.Black, 10, 10, imgWidth - 20, imgHeight - 20);
            }

            // Save the image to a local file.
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Ensure the image file exists before proceeding.
        if (!File.Exists(imagePath))
            throw new Exception($"Image file '{imagePath}' was not created.");

        // -------------------------------------------------
        // Create a simple DOCX template.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("This is a template document.");
        templateDoc.Save(templatePath);

        // Ensure the template file exists.
        if (!File.Exists(templatePath))
            throw new Exception($"Template file '{templatePath}' was not created.");

        // -------------------------------------------------
        // Load the template, insert the image, and save the result.
        // -------------------------------------------------
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph before the image for clarity.
        builder.Writeln("Inserted image below:");

        // Insert the image from the file system.
        builder.InsertImage(imagePath);

        // Save the modified document.
        doc.Save(outputPath);

        // Validate that the output document was created.
        if (!File.Exists(outputPath))
            throw new Exception($"Output file '{outputPath}' was not created.");

        // Optionally, clean up temporary files (commented out to keep files for inspection).
        // File.Delete(imagePath);
        // File.Delete(templatePath);
    }
}

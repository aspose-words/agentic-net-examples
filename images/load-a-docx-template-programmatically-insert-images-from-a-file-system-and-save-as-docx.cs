using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Define file names
        const string templatePath = "template.docx";
        const string imagePath = "input.png";
        const string outputPath = "output.docx";

        // -------------------------------------------------
        // Step 1: Create a sample DOCX template if it does not exist
        // -------------------------------------------------
        if (!File.Exists(templatePath))
        {
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            builder.Writeln("This is a sample template document.");
            builder.Writeln("Below will be an inserted image:");
            templateDoc.Save(templatePath);
        }

        // -------------------------------------------------
        // Step 2: Create a sample image file using Aspose.Drawing
        // -------------------------------------------------
        if (!File.Exists(imagePath))
        {
            const int width = 200;
            const int height = 100;
            Bitmap bitmap = new Bitmap(width, height);
            Graphics graphics = Graphics.FromImage(bitmap);
            // Fill background with white
            graphics.Clear(Color.White);
            // Draw a simple rectangle
            graphics.DrawRectangle(new Pen(Color.Blue, 3), 10, 10, width - 20, height - 20);
            // Save the image
            bitmap.Save(imagePath);
            // Clean up drawing resources
            graphics.Dispose();
            bitmap.Dispose();
        }

        // -------------------------------------------------
        // Step 3: Load the template document
        // -------------------------------------------------
        Document doc = new Document(templatePath);
        DocumentBuilder docBuilder = new DocumentBuilder(doc);

        // Move to the end of the document to insert the image
        docBuilder.MoveToDocumentEnd();

        // -------------------------------------------------
        // Step 4: Insert the image into the document
        // -------------------------------------------------
        if (!File.Exists(imagePath))
            throw new FileNotFoundException("Image file not found.", imagePath);

        docBuilder.InsertParagraph(); // Ensure image starts on a new paragraph
        docBuilder.InsertImage(imagePath);

        // -------------------------------------------------
        // Step 5: Save the resulting document
        // -------------------------------------------------
        doc.Save(outputPath);

        // -------------------------------------------------
        // Validation: Ensure the output file was created
        // -------------------------------------------------
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");

        // Optional: Clean up temporary files (comment out if you want to keep them)
        // File.Delete(templatePath);
        // File.Delete(imagePath);
    }
}

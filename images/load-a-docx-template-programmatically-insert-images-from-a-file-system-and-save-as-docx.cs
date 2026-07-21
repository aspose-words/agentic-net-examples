using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string templatePath = "Template.docx";
        const string imagePath = "input.png";
        const string outputPath = "Output.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a deterministic sample image using Aspose.Drawing.
        // -----------------------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;

        // Ensure any previous files are removed to avoid stale data.
        if (File.Exists(imagePath))
            File.Delete(imagePath);

        // Create a bitmap, draw a solid color, and save it to the file system.
        Bitmap bitmap = new Bitmap(imgWidth, imgHeight);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.CornflowerBlue);
        bitmap.Save(imagePath);
        graphics.Dispose();
        bitmap.Dispose();

        // Verify that the image file was created.
        if (!File.Exists(imagePath))
            throw new Exception($"Failed to create image file '{imagePath}'.");

        // -----------------------------------------------------------------
        // Step 2: Create a simple DOCX template to be loaded later.
        // -----------------------------------------------------------------
        // Remove any existing template.
        if (File.Exists(templatePath))
            File.Delete(templatePath);

        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("Template document. Images will be inserted below:");
        templateDoc.Save(templatePath);

        // Verify that the template file was created.
        if (!File.Exists(templatePath))
            throw new Exception($"Failed to create template file '{templatePath}'.");

        // -----------------------------------------------------------------
        // Step 3: Load the DOCX template.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // Step 4: Insert the previously created image into the document.
        // -----------------------------------------------------------------
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();          // Position cursor at the end.
        builder.InsertParagraph();            // Add a new paragraph.
        builder.InsertImage(imagePath);       // Insert the image inline.

        // -----------------------------------------------------------------
        // Step 5: Save the modified document.
        // -----------------------------------------------------------------
        // Remove any previous output.
        if (File.Exists(outputPath))
            File.Delete(outputPath);

        doc.Save(outputPath);

        // Validate that the output file exists.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to save output document '{outputPath}'.");

        // Optional: Inform that the process completed successfully.
        Console.WriteLine("Document created and saved successfully.");
    }
}

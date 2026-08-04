using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string templatePath = "template.docx";
        const string imagePath1 = "image1.png";
        const string imagePath2 = "image2.png";
        const string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // 1. Create a simple DOCX template (if it does not already exist)
        // -----------------------------------------------------------------
        if (!File.Exists(templatePath))
        {
            Document templateDoc = new Document();
            DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);
            tmplBuilder.Writeln("This is a template document.");
            tmplBuilder.Writeln("Images will be inserted below:");
            templateDoc.Save(templatePath);
        }

        // ---------------------------------------------------------------
        // 2. Create sample images using Aspose.Drawing (deterministic content)
        // ---------------------------------------------------------------
        CreateSampleImage(imagePath1, 200, 100, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(imagePath2, 150, 150, Aspose.Drawing.Color.LightCoral);

        // ---------------------------------------------------------------
        // 3. Load the template document
        // ---------------------------------------------------------------
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move cursor to the end of the document to insert images
        builder.MoveToDocumentEnd();

        // ---------------------------------------------------------------
        // 4. Insert the first image from file system
        // ---------------------------------------------------------------
        Shape shape1 = builder.InsertImage(imagePath1);
        if (!shape1.HasImage)
            throw new InvalidOperationException("First image was not inserted correctly.");

        builder.Writeln(); // Add a line break between images

        // ---------------------------------------------------------------
        // 5. Insert the second image from file system
        // ---------------------------------------------------------------
        Shape shape2 = builder.InsertImage(imagePath2);
        if (!shape2.HasImage)
            throw new InvalidOperationException("Second image was not inserted correctly.");

        // ---------------------------------------------------------------
        // 6. Save the resulting document
        // ---------------------------------------------------------------
        doc.Save(outputPath);

        // ---------------------------------------------------------------
        // 7. Validate that the output file was created
        // ---------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }

    // Helper method to create a deterministic PNG image
    private static void CreateSampleImage(string fileName, int width, int height, Aspose.Drawing.Color background)
    {
        // Ensure any existing file is overwritten
        if (File.Exists(fileName))
            File.Delete(fileName);

        // Create bitmap and graphics objects
        Bitmap bitmap = new Bitmap(width, height);
        Graphics graphics = Graphics.FromImage(bitmap);

        // Fill background
        graphics.Clear(background);

        // Optionally, draw a simple rectangle border
        using (var pen = new Pen(Aspose.Drawing.Color.Black, 2))
        {
            graphics.DrawRectangle(pen, 0, 0, width - 1, height - 1);
        }

        // Save to PNG
        bitmap.Save(fileName);

        // Clean up resources
        graphics.Dispose();
        bitmap.Dispose();
    }
}

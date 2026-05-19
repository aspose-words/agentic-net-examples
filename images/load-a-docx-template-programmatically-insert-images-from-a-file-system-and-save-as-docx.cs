using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // File paths (relative to the executable's working directory)
        string templatePath = "template.docx";
        string imagePath1 = "image1.png";
        string imagePath2 = "image2.png";
        string outputPath = "output.docx";

        // Clean any previous run artifacts
        DeleteIfExists(templatePath);
        DeleteIfExists(imagePath1);
        DeleteIfExists(imagePath2);
        DeleteIfExists(outputPath);

        // 1. Create a simple DOCX template
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("This is a template document.");
        templateBuilder.Writeln("Below are inserted images:");
        templateDoc.Save(templatePath); // Save the template

        // 2. Create sample images using Aspose.Drawing
        CreateSampleImage(imagePath1, 200, 100, Color.LightBlue);
        CreateSampleImage(imagePath2, 150, 150, Color.LightCoral);

        // 3. Load the template document
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move cursor to the end of the document to insert images
        builder.MoveToDocumentEnd();

        // Insert first image
        Shape shape1 = builder.InsertImage(imagePath1);
        shape1.WrapType = WrapType.Inline;

        // Insert a line break between images
        builder.Writeln();

        // Insert second image
        Shape shape2 = builder.InsertImage(imagePath2);
        shape2.WrapType = WrapType.Inline;

        // 4. Save the resulting document
        doc.Save(outputPath);

        // 5. Validate that the output file was created
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create output file: {outputPath}");
    }

    private static void DeleteIfExists(string path)
    {
        if (File.Exists(path))
            File.Delete(path);
    }

    private static void CreateSampleImage(string filePath, int width, int height, Color backgroundColor)
    {
        // Create a bitmap with the specified dimensions
        Bitmap bitmap = new Bitmap(width, height);
        Graphics graphics = Graphics.FromImage(bitmap);
        try
        {
            // Fill the bitmap with a solid background color
            graphics.Clear(backgroundColor);

            // Save the bitmap as a PNG file
            bitmap.Save(filePath);
        }
        finally
        {
            // Ensure resources are released
            graphics.Dispose();
            bitmap.Dispose();
        }
    }
}

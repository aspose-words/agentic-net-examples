using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the output document.
        const string templatePath = "Template.docx";
        const string outputPath = "Modified.docx";

        // -----------------------------------------------------------------
        // Create a simple DOCX template if it does not already exist.
        // -----------------------------------------------------------------
        if (!File.Exists(templatePath))
        {
            Document templateDoc = new Document();
            DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
            templateBuilder.Writeln("This is a template document.");
            templateDoc.Save(templatePath);
        }

        // -----------------------------------------------------------------
        // Load the template document.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // Insert a rectangle shape (inline).
        // -----------------------------------------------------------------
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 150, 80);
        rectangle.FillColor = Color.LightCoral;
        rectangle.Stroke.Color = Color.DarkRed;

        // -----------------------------------------------------------------
        // Insert a floating ellipse shape with custom positioning.
        // -----------------------------------------------------------------
        Shape ellipse = builder.InsertShape(
            ShapeType.Ellipse,
            RelativeHorizontalPosition.Page, 100,   // left distance from page
            RelativeVerticalPosition.Page, 200,     // top distance from page
            120, 120,                               // width, height
            WrapType.None);                         // no text wrapping

        ellipse.FillColor = Color.LightBlue;
        ellipse.Stroke.Color = Color.DarkBlue;

        // -----------------------------------------------------------------
        // Save the modified document.
        // -----------------------------------------------------------------
        doc.Save(outputPath);

        // -----------------------------------------------------------------
        // Validate that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}

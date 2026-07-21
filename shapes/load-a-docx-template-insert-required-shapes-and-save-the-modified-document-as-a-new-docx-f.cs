using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Directory for temporary files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a simple DOCX template and save it.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(artifactsDir, "Template.docx");
        Document templateDoc = new Document(); // blank document with minimal structure
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // 3. Insert a rectangle shape (inline) and set its appearance.
        // -----------------------------------------------------------------
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 150, 80);
        rectangle.FillColor = Color.LightBlue;
        rectangle.Stroke.Color = Color.DarkBlue;
        rectangle.Stroke.Weight = 2.0; // points

        // Add a line break after the rectangle.
        builder.Writeln();

        // -----------------------------------------------------------------
        // 4. Insert a floating text box shape and add some text.
        // -----------------------------------------------------------------
        // Use the overload that creates a floating shape with explicit position.
        Shape textBox = builder.InsertShape(
            ShapeType.TextBox,
            RelativeHorizontalPosition.Page, 100,   // left = 100 points from page left
            RelativeVerticalPosition.Page, 150,    // top  = 150 points from page top
            200,                                   // width
            100,                                   // height
            WrapType.None);                        // no text wrapping

        textBox.FillColor = Color.LightYellow;
        textBox.Stroke.Color = Color.Orange;
        textBox.Stroke.Weight = 1.5;

        // Insert text into the text box.
        // The shape already contains a paragraph (FirstParagraph).
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("Hello from a text box!");

        // -----------------------------------------------------------------
        // 5. Save the modified document.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(artifactsDir, "Modified.docx");
        doc.Save(outputPath);

        // -----------------------------------------------------------------
        // 6. Simple validation – ensure the file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}

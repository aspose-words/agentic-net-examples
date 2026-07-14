using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ShapeInsertionExample
{
    public static void Main()
    {
        // Define file names in the current directory.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple DOCX template and save it.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("This is a template document.");
        templateDoc.Save(templatePath);

        // Verify that the template was created.
        if (!File.Exists(templatePath))
            throw new InvalidOperationException("Failed to create the template document.");

        // -----------------------------------------------------------------
        // 2. Load the template.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // 3. Insert required shapes.
        // -----------------------------------------------------------------
        // Insert an inline rectangle shape.
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        rectangle.Stroke.Color = System.Drawing.Color.Blue;
        rectangle.Fill.Color = System.Drawing.Color.LightGray;

        // Insert a floating text box shape.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
        textBox.WrapType = WrapType.None;
        textBox.Left = 150;   // Position from the left margin (points).
        textBox.Top = 150;    // Position from the top of the page (points).
        textBox.Stroke.Color = System.Drawing.Color.DarkGreen;
        textBox.Fill.Color = System.Drawing.Color.LightYellow;

        // Add text to the text box.
        Paragraph tbParagraph = textBox.FirstParagraph;
        tbParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        tbParagraph.AppendChild(new Run(doc, "Hello from a text box!"));

        // -----------------------------------------------------------------
        // 4. Save the modified document.
        // -----------------------------------------------------------------
        doc.Save(outputPath);

        // -----------------------------------------------------------------
        // 5. Validate that the output file exists.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");

        // The program finishes without waiting for user input.
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define paths for the template and the output document.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "Template.docx");
        string outputPath = Path.Combine(workDir, "Result.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple DOCX template (if it does not already exist).
        // -----------------------------------------------------------------
        if (!File.Exists(templatePath))
        {
            Document templateDoc = new Document();
            DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);
            tmplBuilder.Writeln("This is a template document.");
            templateDoc.Save(templatePath);
        }

        // --------------------------------------------------------------
        // 2. Load the template document.
        // --------------------------------------------------------------
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // --------------------------------------------------------------
        // 3. Insert a floating rectangle shape.
        // --------------------------------------------------------------
        Shape rectangle = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Page, 100,   // left distance from page
            RelativeVerticalPosition.Page, 100,     // top distance from page
            200,                                    // width
            100,                                    // height
            WrapType.None);                         // no text wrapping

        rectangle.FillColor = System.Drawing.Color.LightBlue;
        rectangle.Stroke.Color = System.Drawing.Color.DarkBlue;
        rectangle.StrokeWeight = 2.0;

        // --------------------------------------------------------------
        // 4. Insert an inline text box shape with some text.
        // --------------------------------------------------------------
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 250, 80);
        textBox.FillColor = System.Drawing.Color.LightYellow;
        textBox.Stroke.Color = System.Drawing.Color.Orange;
        textBox.StrokeWeight = 1.5;

        // Add a paragraph and a run inside the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Hello from the text box!");
        para.AppendChild(run);
        textBox.AppendChild(para);

        // --------------------------------------------------------------
        // 5. Save the modified document.
        // --------------------------------------------------------------
        doc.Save(outputPath, SaveFormat.Docx);

        // --------------------------------------------------------------
        // 6. Validate that the output file was created.
        // --------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}

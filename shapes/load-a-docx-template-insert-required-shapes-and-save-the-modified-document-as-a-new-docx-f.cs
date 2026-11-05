using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a simple template document.
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("This is the template document.");
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template document.
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape.
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 150, 80);
        rectangle.FillColor = Color.LightBlue;
        rectangle.StrokeColor = Color.DarkBlue;
        rectangle.StrokeWeight = 2.0f; // Set line width

        // Insert an ellipse shape.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 100, 100);
        ellipse.FillColor = Color.LightGreen;
        ellipse.StrokeColor = Color.DarkGreen;
        ellipse.StrokeWeight = 1.5f; // Set line width

        // Insert a text box shape with some text.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.Width = 200;
        textBox.Height = 100;
        textBox.WrapType = WrapType.Inline;
        textBox.FillColor = Color.LightYellow;
        textBox.StrokeColor = Color.Orange;
        textBox.StrokeWeight = 1.0f; // Set line width

        // Add a paragraph with text inside the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Sample text inside a text box.");
        para.AppendChild(run);
        textBox.AppendChild(para);

        // Append the text box to the document.
        builder.CurrentParagraph.AppendChild(textBox);

        // Save the modified document.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}

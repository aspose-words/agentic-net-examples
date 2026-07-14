using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder for convenience.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a floating text box shape.
        Shape textBox = new Shape(doc, ShapeType.TextBox);

        // Define dimensions (points).
        textBox.Width = 300;   // 300 points wide
        textBox.Height = 100;  // 100 points high

        // Set border (stroke) style.
        textBox.StrokeColor = Color.DarkBlue;          // Border color
        textBox.StrokeWeight = 2.0;                    // Border thickness
        textBox.Stroke.DashStyle = DashStyle.Solid;    // Border dash style

        // Set interior fill color.
        textBox.FillColor = Color.LightYellow;

        // Ensure the shape is floating (no inline wrapping).
        textBox.WrapType = WrapType.None;

        // Add a paragraph inside the text box and some text.
        textBox.AppendChild(new Paragraph(doc));
        Paragraph para = textBox.FirstParagraph;
        para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        Run run = new Run(doc);
        run.Text = "Hello from the text box!";
        para.AppendChild(run);

        // Insert the shape into the document body.
        doc.FirstSection.Body.FirstParagraph.AppendChild(textBox);

        // Define output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TextboxShape.docx");

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}

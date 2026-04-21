using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating textbox shape with specific dimensions.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 50);
        // Make the shape floating by disabling text wrapping around it.
        textBox.WrapType = WrapType.None;

        // Apply border style.
        textBox.Stroke.Color = Color.Blue;
        textBox.Stroke.Weight = 2.0;          // Correct property for line width.
        textBox.Stroke.DashStyle = DashStyle.Dash;

        // Apply interior fill formatting.
        textBox.FillColor = Color.LightYellow;
        textBox.Filled = true;

        // Add a paragraph with text inside the textbox.
        textBox.AppendChild(new Paragraph(doc));
        builder.MoveTo(textBox.LastParagraph);
        builder.Font.Size = 12;
        builder.Write("Hello, Aspose.Words textbox!");

        // Save the document to a local file.
        string outputPath = "TextBoxShape.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}

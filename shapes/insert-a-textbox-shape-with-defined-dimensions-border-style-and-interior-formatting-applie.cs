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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a textbox shape with specific dimensions.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);

        // Apply border (stroke) formatting.
        textBox.StrokeColor = Color.Blue;          // Border color.
        textBox.StrokeWeight = 2.0;                // Border thickness (points).
        textBox.Stroke.DashStyle = DashStyle.Solid; // Border style.

        // Apply interior (fill) formatting.
        textBox.FillColor = Color.LightYellow;     // Background color.

        // Optionally add some text inside the textbox.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Sample textbox content");
        para.AppendChild(run);
        textBox.AppendChild(para);

        // Save the document to disk.
        string outputPath = "TextboxShape.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}

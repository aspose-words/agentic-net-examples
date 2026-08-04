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

        // Insert a floating text box with specific dimensions (width: 200 pt, height: 100 pt).
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
        // Make the shape floating so that we can set its position and wrapping.
        textBox.WrapType = WrapType.None;

        // Apply border (stroke) formatting.
        textBox.StrokeColor = Color.DarkBlue;               // Border color.
        textBox.StrokeWeight = 2.0;                         // Border thickness (points).
        textBox.Stroke.DashStyle = DashStyle.Dash;          // Dashed border.

        // Apply interior (fill) formatting.
        textBox.FillColor = Color.LightYellow;              // Background color of the text box.

        // Optional: adjust internal margins of the text box (in points).
        textBox.TextBox.InternalMarginTop = 5;
        textBox.TextBox.InternalMarginBottom = 5;
        textBox.TextBox.InternalMarginLeft = 5;
        textBox.TextBox.InternalMarginRight = 5;

        // Add some text inside the text box.
        builder.MoveTo(textBox.LastParagraph);
        builder.Font.Size = 12;
        builder.Write("This is a sample text inside the formatted text box.");

        // Save the document to the local file system.
        string outputPath = "TextboxShape.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create the output file: {outputPath}");
    }
}

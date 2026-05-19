using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating text box shape with a specific size.
        Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 200, 100);
        // Ensure the shape is floating so we can control its position.
        textBoxShape.WrapType = WrapType.None;

        // Set the anchor position to be relative to the page margins.
        textBoxShape.RelativeVerticalPosition = RelativeVerticalPosition.Margin;
        textBoxShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;

        // Position the text box 50 points from the top and left margins.
        textBoxShape.Top = 50;
        textBoxShape.Left = 50;

        // Add some text inside the text box.
        builder.MoveTo(textBoxShape.LastParagraph);
        builder.Write("Hello from a text box!");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TextBoxAnchor.docx");
        doc.Save(outputPath);
    }
}

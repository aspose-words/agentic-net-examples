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
        // Ensure the shape is floating rather than inline.
        textBoxShape.WrapType = WrapType.None;

        // Set the anchor position to be relative to the page margins.
        textBoxShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
        textBoxShape.RelativeVerticalPosition = RelativeVerticalPosition.Margin;

        // Optional: set offsets from the margins (0 points in this example).
        textBoxShape.Left = 0;
        textBoxShape.Top = 0;

        // Add some text inside the text box.
        builder.MoveTo(textBoxShape.LastParagraph);
        builder.Write("Hello inside text box.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TextBoxRelativeToMargin.docx");
        doc.Save(outputPath);
    }
}

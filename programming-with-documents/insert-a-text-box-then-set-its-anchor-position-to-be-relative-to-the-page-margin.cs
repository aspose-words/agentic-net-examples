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

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating text box shape with a specific size.
        Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 200, 100);
        // Ensure the shape is floating (not inline) so that positioning properties apply.
        textBoxShape.WrapType = WrapType.None;

        // Set the anchor position to be relative to the page margins.
        textBoxShape.RelativeVerticalPosition = RelativeVerticalPosition.Margin;
        textBoxShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;

        // Optionally add some text inside the text box.
        builder.MoveTo(textBoxShape.LastParagraph);
        builder.Write("Text inside the textbox.");

        // Determine an output path in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TextBoxRelativeToMargin.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}

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

        // Insert a floating text box shape.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
        // Make the shape floating so that positioning properties take effect.
        textBox.WrapType = WrapType.None;

        // Anchor the text box relative to the page margins.
        textBox.RelativeVerticalPosition = RelativeVerticalPosition.Margin;
        textBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;

        // Add some text inside the text box.
        builder.MoveTo(textBox.LastParagraph);
        builder.Write("This text box is anchored to the page margin.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TextBoxAnchorMargin.docx");
        doc.Save(outputPath);
    }
}

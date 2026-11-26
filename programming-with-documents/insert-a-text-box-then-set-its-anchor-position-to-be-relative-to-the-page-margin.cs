using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text box shape with specified size.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);

        // Set the anchor position of the text box relative to the page margin.
        textBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
        textBox.RelativeVerticalPosition = RelativeVerticalPosition.Margin;

        // Optionally align the text box to the top-left corner of the margin.
        textBox.HorizontalAlignment = HorizontalAlignment.Left;
        textBox.VerticalAlignment = VerticalAlignment.Top;
        textBox.Left = 0;
        textBox.Top = 0;

        // Add some text inside the text box.
        builder.MoveTo(textBox.FirstParagraph);
        builder.Writeln("Hello, TextBox!");

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}

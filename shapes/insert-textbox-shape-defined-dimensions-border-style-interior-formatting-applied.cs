using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating textbox shape with specific dimensions (200x100 points).
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);

        // Configure the shape to float and set its position on the page.
        textBox.WrapType = WrapType.None; // No text wrapping around the shape.
        textBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        textBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        textBox.Left = 100; // Distance from the left edge of the page (points).
        textBox.Top = 150;  // Distance from the top edge of the page (points).

        // Apply border (stroke) formatting.
        textBox.Stroked = true;               // Enable the border.
        textBox.StrokeColor = Color.DarkBlue; // Border color.
        textBox.StrokeWeight = 2.0;           // Border thickness in points.

        // Apply interior (fill) formatting.
        textBox.Filled = true;                // Enable filling.
        textBox.FillColor = Color.LightYellow; // Fill color.

        // Set internal margins for the text inside the textbox.
        TextBox tb = textBox.TextBox;
        tb.InternalMarginTop = 10;
        tb.InternalMarginBottom = 10;
        tb.InternalMarginLeft = 10;
        tb.InternalMarginRight = 10;

        // Add a paragraph inside the textbox and write some text.
        builder.MoveTo(textBox.FirstParagraph);
        builder.Font.Size = 12;
        builder.Font.Name = "Arial";
        builder.Writeln("This is a custom textbox with border and fill.");

        // Save the document.
        doc.Save("TextboxExample.docx");
    }
}

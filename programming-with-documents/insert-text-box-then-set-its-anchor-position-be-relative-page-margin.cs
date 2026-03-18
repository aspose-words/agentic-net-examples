using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating text box. Use the overload that allows specifying
        // relative positioning and wrap type.
        Shape textBox = builder.InsertShape(
            ShapeType.TextBox,                     // shape type
            RelativeHorizontalPosition.Margin,     // horizontal position relative to page margin
            100,                                   // left offset (points) from the margin
            RelativeVerticalPosition.Margin,       // vertical position relative to page margin
            100,                                   // top offset (points) from the margin
            200,                                   // width (points)
            100,                                   // height (points)
            WrapType.None);                        // no text wrapping

        // Optional: set additional properties, e.g., no border.
        textBox.WrapType = WrapType.None;

        // Add some text inside the text box.
        builder.MoveTo(textBox.LastParagraph);
        builder.Write("This is a floating text box anchored to the page margin.");

        // Save the document.
        doc.Save("TextBoxAnchorMargin.docx");
    }
}

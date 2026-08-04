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

        // Insert a floating text box shape.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.WrapType = WrapType.None;
        textBox.Width = 200;   // Width in points.
        textBox.Height = 50;   // Height in points.

        // Position the original text box at (50, 50) points from the page.
        textBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        textBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        textBox.Left = 50;
        textBox.Top = 50;

        // Add a paragraph and a run of text inside the text box.
        textBox.AppendChild(new Paragraph(doc));
        Paragraph para = textBox.FirstParagraph;
        Run run = new Run(doc, "Original TextBox");
        para.AppendChild(run);

        // Insert the original text box into the document body.
        builder.InsertParagraph();                     // Ensure we have a paragraph to work with.
        builder.CurrentParagraph.AppendChild(textBox); // Add the shape to the current paragraph.

        // Clone the text box.
        Shape clonedBox = (Shape)textBox.Clone(true);

        // Place the cloned box at a new absolute position, e.g., (300, 400) points from the page.
        clonedBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        clonedBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        clonedBox.Left = 300;
        clonedBox.Top = 400;

        // Insert the cloned text box into the document.
        builder.CurrentParagraph.AppendChild(clonedBox);

        // Save the resulting document.
        doc.Save("DuplicatedTextBox.docx");
    }
}

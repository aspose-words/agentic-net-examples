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
        textBox.WrapType = WrapType.None;          // Make it free‑floating.
        textBox.Width = 200;                       // Width in points.
        textBox.Height = 50;                       // Height in points.
        textBox.Left = 50;                         // Initial horizontal position.
        textBox.Top = 50;                          // Initial vertical position.
        textBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        textBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;

        // Add a paragraph with some text inside the text box.
        textBox.AppendChild(new Paragraph(doc));
        Paragraph para = textBox.FirstParagraph;
        para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        Run run = new Run(doc, "Original TextBox");
        para.AppendChild(run);

        // Place the original text box into the document.
        builder.InsertParagraph();                 // Ensure we have a paragraph to host the shape.
        builder.CurrentParagraph.AppendChild(textBox);

        // Clone the text box node (deep copy).
        Shape clonedBox = (Shape)textBox.Clone(true);

        // Position the cloned box at an absolute location on the page.
        clonedBox.Left = 300;                      // Horizontal position in points.
        clonedBox.Top = 200;                       // Vertical position in points.
        clonedBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        clonedBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;

        // Optionally change the text inside the cloned box.
        clonedBox.FirstParagraph.Runs[0].Text = "Cloned TextBox";

        // Insert the cloned text box after the original one.
        builder.CurrentParagraph.AppendChild(clonedBox);

        // Save the document to a file.
        doc.Save("DuplicatedTextBox.docx");
    }
}

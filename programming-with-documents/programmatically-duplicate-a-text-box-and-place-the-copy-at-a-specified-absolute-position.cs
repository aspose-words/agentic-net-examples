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

        // Insert a text box shape.
        Shape originalBox = new Shape(doc, ShapeType.TextBox);
        originalBox.Width = 200;
        originalBox.Height = 100;
        originalBox.WrapType = WrapType.None;
        originalBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        originalBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        originalBox.Left = 50;   // initial position
        originalBox.Top = 50;

        // Add a paragraph with some text inside the text box.
        originalBox.AppendChild(new Paragraph(doc));
        originalBox.FirstParagraph.AppendChild(new Run(doc, "Original Text Box"));

        // Insert the original shape into the document.
        builder.InsertNode(originalBox);

        // Clone the original text box.
        Shape clonedBox = (Shape)originalBox.Clone(true);
        // Set the absolute position for the cloned box.
        clonedBox.Left = 300; // points from the left edge of the page
        clonedBox.Top = 200;  // points from the top edge of the page

        // Insert the cloned shape into the document.
        builder.InsertNode(clonedBox);

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DuplicatedTextBox.docx");
        doc.Save(outputPath);
    }
}

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

        // Insert the first text box and add an empty paragraph inside it.
        Shape shape1 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        shape1.AppendChild(new Paragraph(doc));
        TextBox textBox1 = shape1.TextBox;

        // Insert the second text box and add an empty paragraph inside it.
        Shape shape2 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        shape2.AppendChild(new Paragraph(doc));
        TextBox textBox2 = shape2.TextBox;

        // Link the first text box to the second if the link is valid.
        if (textBox1.IsValidLinkTarget(textBox2))
        {
            textBox1.Next = textBox2;
        }

        // Add a long paragraph to the first text box to cause overflow.
        builder.MoveTo(shape1.FirstParagraph);
        builder.Writeln(
            "This is a long text that will overflow from the first text box to the linked second text box. " +
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
            "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
            "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat."
        );

        // Optionally add some text directly to the second text box.
        builder.MoveTo(shape2.FirstParagraph);
        builder.Writeln("Additional text in the second box.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "LinkedTextBoxes.docx");
        doc.Save(outputPath);
    }
}

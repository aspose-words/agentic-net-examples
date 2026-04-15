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

        // Insert the first text box.
        Shape textBoxShape1 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        TextBox textBox1 = textBoxShape1.TextBox;

        // Add a long paragraph inside the first text box.
        builder.MoveTo(textBoxShape1.LastParagraph);
        builder.Writeln(
            "This is a long text that will overflow from the first text box. " +
            "It contains enough words to exceed the size of the box and demonstrate " +
            "automatic flow to the linked text box. " +
            "The quick brown fox jumps over the lazy dog. " +
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
            "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // Insert the second text box (the target of the link).
        builder.Writeln(); // Add a line break between the boxes.
        Shape textBoxShape2 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        TextBox textBox2 = textBoxShape2.TextBox;

        // Link the first text box to the second if the link is valid.
        if (textBox1.IsValidLinkTarget(textBox2))
        {
            textBox1.Next = textBox2;
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "LinkedTextBoxes.docx");
        doc.Save(outputPath);
    }
}

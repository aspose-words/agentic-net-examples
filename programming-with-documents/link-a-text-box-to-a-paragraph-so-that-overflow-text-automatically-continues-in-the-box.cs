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

        // Insert the first floating text box.
        Shape textBox1 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        textBox1.WrapType = WrapType.None; // Prevent text wrapping around the shape.

        // Add a paragraph inside the first text box and fill it with long text.
        Paragraph para1 = textBox1.FirstParagraph;
        para1.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
        Run longRun = new Run(doc, new string('A', 1000)); // Text that will overflow.
        para1.AppendChild(longRun);

        // Insert the second text box that will receive the overflow text.
        Shape textBox2 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        textBox2.WrapType = WrapType.None;

        // Link the first text box to the second one.
        textBox1.TextBox.Next = textBox2.TextBox;

        // Save the resulting document.
        doc.Save("LinkedTextBox.docx");
    }
}

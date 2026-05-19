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

        // Insert three text boxes into the document.
        Shape shape1 = builder.InsertShape(ShapeType.TextBox, 100, 100);
        TextBox textBox1 = shape1.TextBox;
        builder.Writeln();

        Shape shape2 = builder.InsertShape(ShapeType.TextBox, 100, 100);
        TextBox textBox2 = shape2.TextBox;
        builder.Writeln();

        Shape shape3 = builder.InsertShape(ShapeType.TextBox, 100, 100);
        TextBox textBox3 = shape3.TextBox;
        builder.Writeln();

        // Link the text boxes to form a sequence: 1 -> 2 -> 3.
        if (textBox1.IsValidLinkTarget(textBox2))
            textBox1.Next = textBox2;

        if (textBox2.IsValidLinkTarget(textBox3))
            textBox2.Next = textBox3;

        // Determine if the first text box is the head of the linked sequence.
        // A head has no previous textbox and has a next textbox.
        bool isHead = textBox1.Next != null && textBox1.Previous == null;

        Console.WriteLine(isHead
            ? "The first TextBox is the head of the linked sequence."
            : "The first TextBox is not the head of the linked sequence.");

        // Save the resulting document.
        doc.Save("LinkedTextBoxes.docx");
    }
}

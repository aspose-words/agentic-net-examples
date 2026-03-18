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

        // Insert three empty text boxes.
        Shape shape1 = builder.InsertShape(ShapeType.TextBox, 100, 100);
        TextBox textBox1 = shape1.TextBox;
        builder.Writeln(); // separate paragraphs

        Shape shape2 = builder.InsertShape(ShapeType.TextBox, 100, 100);
        TextBox textBox2 = shape2.TextBox;
        builder.Writeln();

        Shape shape3 = builder.InsertShape(ShapeType.TextBox, 100, 100);
        TextBox textBox3 = shape3.TextBox;

        // Link the text boxes: 1 -> 2 -> 3.
        if (textBox1.IsValidLinkTarget(textBox2))
            textBox1.Next = textBox2;

        if (textBox2.IsValidLinkTarget(textBox3))
            textBox2.Next = textBox3;

        // Determine whether a given TextBox is the head of the linked sequence.
        // A head has a next link but no previous link.
        bool isHead = textBox1.Next != null && textBox1.Previous == null;

        Console.WriteLine(isHead
            ? "TextBox1 is the head of the sequence."
            : "TextBox1 is NOT the head of the sequence.");

        // Save the document (optional, just to demonstrate lifecycle usage).
        doc.Save("LinkedTextBoxes.docx");
    }
}

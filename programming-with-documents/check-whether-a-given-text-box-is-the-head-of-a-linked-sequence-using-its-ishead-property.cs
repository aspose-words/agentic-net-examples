using System;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three text boxes.
        Shape shape1 = builder.InsertShape(ShapeType.TextBox, 100, 100);
        TextBox textBox1 = shape1.TextBox;
        builder.Writeln();

        Shape shape2 = builder.InsertShape(ShapeType.TextBox, 100, 100);
        TextBox textBox2 = shape2.TextBox;
        builder.Writeln();

        Shape shape3 = builder.InsertShape(ShapeType.TextBox, 100, 100);
        TextBox textBox3 = shape3.TextBox;
        builder.Writeln();

        // Link the text boxes: 1 -> 2 -> 3.
        if (textBox1.IsValidLinkTarget(textBox2))
            textBox1.Next = textBox2;

        if (textBox2.IsValidLinkTarget(textBox3))
            textBox2.Next = textBox3;

        // Determine whether the first text box is the head of the linked sequence.
        // Prefer the IsHead property if it exists; otherwise fall back to manual logic.
        bool isHead;
        PropertyInfo isHeadProp = typeof(TextBox).GetProperty("IsHead");
        if (isHeadProp != null)
        {
            isHead = (bool)isHeadProp.GetValue(textBox1);
        }
        else
        {
            // A head has no previous link but has a next link.
            isHead = textBox1.Previous == null && textBox1.Next != null;
        }

        Console.WriteLine($"TextBox 1 is head of the sequence: {isHead}");

        // Save the document (optional, just to demonstrate a complete workflow).
        doc.Save("LinkedTextBoxes.docx");
    }
}

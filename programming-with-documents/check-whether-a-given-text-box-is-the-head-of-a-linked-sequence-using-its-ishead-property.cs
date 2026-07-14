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
        bool isHead = textBox1.Next != null && textBox1.Previous == null;
        Console.WriteLine($"TextBox 1 is head of the sequence: {isHead}");

        // Optionally, verify other boxes.
        bool isMiddle = textBox2.Next != null && textBox2.Previous != null;
        bool isTail = textBox3.Next == null && textBox3.Previous != null;
        Console.WriteLine($"TextBox 2 is middle of the sequence: {isMiddle}");
        Console.WriteLine($"TextBox 3 is tail of the sequence: {isTail}");

        // Save the document.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "LinkedTextBoxes.docx");
        doc.Save(outputPath);
    }
}

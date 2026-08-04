using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define a folder for output files and ensure it exists.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three text boxes into the document.
        Shape shape1 = builder.InsertShape(ShapeType.TextBox, 100, 50);
        TextBox textBox1 = shape1.TextBox;
        builder.Writeln();

        Shape shape2 = builder.InsertShape(ShapeType.TextBox, 100, 50);
        TextBox textBox2 = shape2.TextBox;
        builder.Writeln();

        Shape shape3 = builder.InsertShape(ShapeType.TextBox, 100, 50);
        TextBox textBox3 = shape3.TextBox;
        builder.Writeln();

        // Link the text boxes into a sequence (1 -> 2 -> 3) if linking is allowed.
        if (textBox1.IsValidLinkTarget(textBox2))
            textBox1.Next = textBox2;

        if (textBox2.IsValidLinkTarget(textBox3))
            textBox2.Next = textBox3;

        // Determine whether each text box is the head of the linked sequence.
        // A text box is the head when it has a next link but no previous link.
        bool isHead1 = textBox1.Next != null && textBox1.Previous == null;
        bool isHead2 = textBox2.Next != null && textBox2.Previous == null;
        bool isHead3 = textBox3.Next != null && textBox3.Previous == null;

        Console.WriteLine($"TextBox 1 IsHead: {isHead1}");
        Console.WriteLine($"TextBox 2 IsHead: {isHead2}");
        Console.WriteLine($"TextBox 3 IsHead: {isHead3}");

        // Save the document so the result can be inspected if needed.
        doc.Save(Path.Combine(artifactsDir, "LinkedTextBoxes.docx"));
    }
}

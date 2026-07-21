using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three empty text boxes.
        Shape shape1 = builder.InsertShape(ShapeType.TextBox, 100, 50);
        TextBox textBox1 = shape1.TextBox;
        builder.Writeln();

        Shape shape2 = builder.InsertShape(ShapeType.TextBox, 100, 50);
        TextBox textBox2 = shape2.TextBox;
        builder.Writeln();

        Shape shape3 = builder.InsertShape(ShapeType.TextBox, 100, 50);
        TextBox textBox3 = shape3.TextBox;
        builder.Writeln();

        // Link the text boxes into a sequence: 1 -> 2 -> 3.
        if (textBox1.IsValidLinkTarget(textBox2))
            textBox1.Next = textBox2;

        if (textBox2.IsValidLinkTarget(textBox3))
            textBox2.Next = textBox3;

        // Determine whether the first text box is the head of the linked sequence.
        // A head has a Next link but no Previous link.
        bool isHead = textBox1.Next != null && textBox1.Previous == null;
        Console.WriteLine($"TextBox 1 is head of the sequence: {isHead}");

        // Save the document.
        string outputPath = Path.Combine(artifactsDir, "LinkedTextBoxes.docx");
        doc.Save(outputPath);
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define output folder and ensure it exists.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert four text boxes.
        Shape shape1 = builder.InsertShape(ShapeType.TextBox, 100, 100);
        TextBox textBox1 = shape1.TextBox;
        builder.Writeln();

        Shape shape2 = builder.InsertShape(ShapeType.TextBox, 100, 100);
        TextBox textBox2 = shape2.TextBox;
        builder.Writeln();

        Shape shape3 = builder.InsertShape(ShapeType.TextBox, 100, 100);
        TextBox textBox3 = shape3.TextBox;
        builder.Writeln();

        Shape shape4 = builder.InsertShape(ShapeType.TextBox, 100, 100);
        TextBox textBox4 = shape4.TextBox;

        // Link the first three text boxes: 1 -> 2 -> 3.
        if (textBox1.IsValidLinkTarget(textBox2))
            textBox1.Next = textBox2;

        if (textBox2.IsValidLinkTarget(textBox3))
            textBox2.Next = textBox3;

        // Break the forward link of the middle text box (textBox2) so that it no longer links to textBox3.
        // The Previous property of textBox3 points to textBox2.
        if (textBox3.Previous != null)
            textBox3.Previous.BreakForwardLink();

        // Save the resulting document.
        doc.Save(Path.Combine(artifactsDir, "BreakForwardLink.docx"));
    }
}

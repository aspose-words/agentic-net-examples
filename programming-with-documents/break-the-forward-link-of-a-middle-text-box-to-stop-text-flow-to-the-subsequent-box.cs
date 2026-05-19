using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class BreakForwardLinkExample
{
    public static void Main()
    {
        // Define output folder and ensure it exists.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first text box.
        Shape shape1 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        TextBox textBox1 = shape1.TextBox;
        builder.Writeln(); // Move cursor after the shape.

        // Insert the second text box (the middle one we will break the link from).
        Shape shape2 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        TextBox textBox2 = shape2.TextBox;
        builder.Writeln();

        // Insert the third text box.
        Shape shape3 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        TextBox textBox3 = shape3.TextBox;
        builder.Writeln();

        // Insert a fourth text box to demonstrate that text flow stops after breaking the link.
        Shape shape4 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        TextBox textBox4 = shape4.TextBox;
        // Move the builder inside the fourth text box and add some text.
        builder.MoveTo(shape4.LastParagraph);
        builder.Write("This box is not linked.");

        // Link the first three text boxes together (1 -> 2 -> 3).
        if (textBox1.IsValidLinkTarget(textBox2))
            textBox1.Next = textBox2;

        if (textBox2.IsValidLinkTarget(textBox3))
            textBox2.Next = textBox3;

        // Break the forward link of the middle text box (textBox2) so that it no longer links to textBox3.
        // The Previous property of textBox3 points to textBox2.
        if (textBox3.Previous != null)
            textBox3.Previous.BreakForwardLink();

        // Save the resulting document.
        string outputPath = Path.Combine(artifactsDir, "BreakForwardLink.docx");
        doc.Save(outputPath);
    }
}

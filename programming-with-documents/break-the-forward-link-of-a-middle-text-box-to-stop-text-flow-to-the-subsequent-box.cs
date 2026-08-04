using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert four text boxes into the document.
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

        // Link the first three text boxes: 1 → 2 → 3.
        if (textBox1.IsValidLinkTarget(textBox2))
            textBox1.Next = textBox2;

        if (textBox2.IsValidLinkTarget(textBox3))
            textBox2.Next = textBox3;

        // Break the forward link of the middle text box (textBox2).
        // After this call, textBox2.Next becomes null and textBox3.Previous becomes null.
        textBox2.BreakForwardLink();

        // Save the resulting document.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "BreakForwardLink.docx");
        doc.Save(outputPath);
    }
}

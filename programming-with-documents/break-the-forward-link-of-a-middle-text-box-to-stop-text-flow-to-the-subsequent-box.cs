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

        // Insert four text boxes.
        Shape shape1 = builder.InsertShape(ShapeType.TextBox, 150, 100);
        TextBox textBox1 = shape1.TextBox;
        builder.Writeln();

        Shape shape2 = builder.InsertShape(ShapeType.TextBox, 150, 100);
        TextBox textBox2 = shape2.TextBox;
        builder.Writeln();

        Shape shape3 = builder.InsertShape(ShapeType.TextBox, 150, 100);
        TextBox textBox3 = shape3.TextBox;
        builder.Writeln();

        Shape shape4 = builder.InsertShape(ShapeType.TextBox, 150, 100);
        TextBox textBox4 = shape4.TextBox;

        // Link the first three text boxes (1 -> 2 -> 3).
        if (textBox1.IsValidLinkTarget(textBox2))
            textBox1.Next = textBox2;

        if (textBox2.IsValidLinkTarget(textBox3))
            textBox2.Next = textBox3;

        // Verify that the fourth box is not a valid link target for the third (it must be empty).
        // This is just for demonstration; no assertion library is used.
        bool canLinkToFourth = textBox3.IsValidLinkTarget(textBox4);

        // Write some text into the fourth text box to make it non‑empty.
        builder.MoveTo(shape4.LastParagraph);
        builder.Write("Fourth box content");

        // After adding content, the third box can no longer link to the fourth.
        // Break the forward link between the second and third text boxes.
        // This stops the text flow from the second box to the third.
        textBox2.BreakForwardLink();

        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Save the document.
        doc.Save(Path.Combine(artifactsDir, "BreakForwardLink.docx"));
    }
}

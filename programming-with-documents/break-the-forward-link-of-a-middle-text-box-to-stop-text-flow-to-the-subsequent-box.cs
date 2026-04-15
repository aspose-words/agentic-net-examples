using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace BreakForwardLinkExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert four text boxes into the document.
            Shape shape1 = builder.InsertShape(ShapeType.TextBox, 100, 100);
            TextBox textBox1 = shape1.TextBox;
            builder.Writeln(); // Move cursor after the first box.

            Shape shape2 = builder.InsertShape(ShapeType.TextBox, 100, 100);
            TextBox textBox2 = shape2.TextBox;
            builder.Writeln();

            Shape shape3 = builder.InsertShape(ShapeType.TextBox, 100, 100);
            TextBox textBox3 = shape3.TextBox;
            builder.Writeln();

            Shape shape4 = builder.InsertShape(ShapeType.TextBox, 100, 100);
            TextBox textBox4 = shape4.TextBox;
            builder.Writeln();

            // Link the text boxes in a sequence: 1 → 2 → 3 → 4.
            if (textBox1.IsValidLinkTarget(textBox2))
                textBox1.Next = textBox2;

            if (textBox2.IsValidLinkTarget(textBox3))
                textBox2.Next = textBox3;

            if (textBox3.IsValidLinkTarget(textBox4))
                textBox3.Next = textBox4;

            // Break the forward link of the middle text box (textBox2).
            // This stops the flow of text from textBox2 to textBox3.
            textBox2.BreakForwardLink();

            // Prepare the output folder.
            string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
            Directory.CreateDirectory(artifactsDir);

            // Save the resulting document.
            string outputPath = Path.Combine(artifactsDir, "BreakForwardLink.docx");
            doc.Save(outputPath);
        }
    }
}

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

        // Insert the first text box.
        Shape shape1 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        TextBox textBox1 = shape1.TextBox;

        // Insert the second text box.
        Shape shape2 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        TextBox textBox2 = shape2.TextBox;

        // Link the first text box to the second so that overflow text continues automatically.
        if (textBox1.IsValidLinkTarget(textBox2))
        {
            textBox1.Next = textBox2;
        }

        // Move the cursor inside the first text box and write a long paragraph.
        builder.MoveTo(shape1.FirstParagraph);
        builder.Writeln("This is a long piece of text that will not fit entirely within the first text box. " +
                        "It should automatically continue into the second linked text box, demonstrating " +
                        "the sequential linking of text boxes using Aspose.Words. " +
                        "The quick brown fox jumps over the lazy dog. " +
                        "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt " +
                        "ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation " +
                        "ullamco laboris nisi ut aliquip ex ea commodo consequat.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "LinkedTextBoxes.docx");
        doc.Save(outputPath);
    }
}

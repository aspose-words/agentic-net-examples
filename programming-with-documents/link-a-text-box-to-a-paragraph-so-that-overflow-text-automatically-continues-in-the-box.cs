using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a regular paragraph before the text boxes.
        builder.Writeln("Paragraph before linked text boxes.");

        // Insert the first text box.
        Shape textBox1 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        // Make the shape grow to fit its contents.
        textBox1.TextBox.FitShapeToText = true;
        // Prevent internal wrapping so the shape can expand.
        textBox1.TextBox.TextBoxWrapMode = TextBoxWrapMode.None;

        // Insert the second text box that will receive overflow text.
        Shape textBox2 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        textBox2.TextBox.FitShapeToText = true;
        textBox2.TextBox.TextBoxWrapMode = TextBoxWrapMode.None;

        // Link the first text box to the second one.
        if (textBox1.TextBox.IsValidLinkTarget(textBox2.TextBox))
            textBox1.TextBox.Next = textBox2.TextBox;

        // Move the cursor inside the first text box and write a long text.
        builder.MoveTo(textBox1.LastParagraph);
        builder.Font.Size = 12;
        builder.Write("This is a long piece of text that will not fit entirely within the first text box. ");
        builder.Write("It should automatically continue in the linked text box. ");
        builder.Write("Adding more sentences to ensure overflow occurs. ");
        builder.Write("The quick brown fox jumps over the lazy dog. ");
        builder.Write("Lorem ipsum dolor sit amet, consectetur adipiscing elit. ");

        // Add a paragraph after the linked text boxes.
        builder.Writeln();
        builder.Writeln("Paragraph after linked text boxes.");

        // Save the resulting document.
        doc.Save("LinkedTextBox.docx");
    }
}

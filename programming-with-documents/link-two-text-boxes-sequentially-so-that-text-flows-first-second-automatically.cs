using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first text box (200x100 points).
        Shape textBoxShape1 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        TextBox textBox1 = textBoxShape1.TextBox;

        // Write some text that will overflow the first box.
        builder.MoveTo(textBoxShape1.LastParagraph);
        builder.Writeln("This is a long piece of text that will not fit entirely inside the first text box. ");

        // Insert the second text box that will receive overflow text.
        Shape textBoxShape2 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        TextBox textBox2 = textBoxShape2.TextBox;

        // The second box must contain at least one empty paragraph before linking.
        builder.Writeln();

        // Link the first text box to the second if the link is valid.
        if (textBox1.IsValidLinkTarget(textBox2))
            textBox1.Next = textBox2;

        // Add more text to the first box; overflow will automatically continue in the second box.
        builder.MoveTo(textBoxShape1.LastParagraph);
        builder.Writeln("Additional text that should flow into the linked textbox automatically.");

        // Save the resulting document.
        doc.Save("LinkedTextBoxes.docx");
    }
}

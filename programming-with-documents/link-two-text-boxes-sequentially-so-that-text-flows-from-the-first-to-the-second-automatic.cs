using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a DocumentBuilder to edit it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first floating text box.
        Shape shape1 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        TextBox textBox1 = shape1.TextBox;

        // Insert the second floating text box.
        Shape shape2 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        TextBox textBox2 = shape2.TextBox;

        // Link the first text box to the second one so that overflow text flows automatically.
        if (textBox1.IsValidLinkTarget(textBox2))
        {
            textBox1.Next = textBox2;
        }

        // Add a long piece of text to the first text box.
        // When the text exceeds the bounds of the first box, it will continue in the linked second box.
        builder.MoveTo(shape1.FirstParagraph);
        builder.Write("This is a long piece of text that will not fit entirely within the first text box. " +
                      "Aspose.Words will automatically continue the overflow into the second linked text box, " +
                      "demonstrating sequential text flow between linked text boxes.");

        // Save the document to the local file system.
        doc.Save("LinkedTextBoxes.docx");
    }
}

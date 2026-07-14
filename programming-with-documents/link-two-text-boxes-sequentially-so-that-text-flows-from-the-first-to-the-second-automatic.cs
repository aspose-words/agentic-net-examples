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

        // Insert the first text box.
        Shape shape1 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        TextBox textBox1 = shape1.TextBox;

        // Insert the second text box.
        Shape shape2 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        TextBox textBox2 = shape2.TextBox;

        // Link the first text box to the second if the link is valid.
        if (textBox1.IsValidLinkTarget(textBox2))
        {
            textBox1.Next = textBox2;
        }

        // Move the cursor into the first text box and write a long text.
        // The overflow will automatically continue into the linked second text box.
        builder.MoveTo(shape1.LastParagraph);
        builder.Write("This is a long text that will automatically flow from the first text box to the linked second text box. ");
        builder.Write("Adding more content to ensure overflow. ");
        builder.Write("Even more text to demonstrate the linking functionality. ");

        // Save the document.
        doc.Save("LinkedTextBoxes.docx");
    }
}

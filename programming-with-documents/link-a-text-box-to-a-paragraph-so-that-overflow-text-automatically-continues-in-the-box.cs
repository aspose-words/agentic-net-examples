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
        Shape textBoxShape1 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        TextBox textBox1 = textBoxShape1.TextBox;

        // Insert the second text box that will receive overflow text.
        Shape textBoxShape2 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        TextBox textBox2 = textBoxShape2.TextBox;

        // Link the first text box to the second one.
        textBox1.Next = textBox2;

        // Move the cursor inside the first text box and write a long paragraph.
        builder.MoveTo(textBoxShape1.LastParagraph);
        string longText = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                          "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                          "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris " +
                          "nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in " +
                          "reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla " +
                          "pariatur. Excepteur sint occaecat cupidatat non proident, sunt in " +
                          "culpa qui officia deserunt mollit anim id est laborum. ";
        // Repeat to ensure overflow.
        for (int i = 0; i < 5; i++)
        {
            builder.Write(longText);
        }

        // Save the document.
        doc.Save("LinkedTextBoxes.docx");
    }
}

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

        // Long text that will not fit into a single text box.
        string longText = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                          "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                          "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris " +
                          "nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in " +
                          "reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. " +
                          "Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia " +
                          "deserunt mollit anim id est laborum. ";

        // Insert the first text box.
        Shape textBox1 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        // Prevent the shape from automatically resizing to fit the text.
        textBox1.TextBox.FitShapeToText = false;

        // Insert the second text box (will receive overflow text).
        builder.Writeln(); // Move cursor to a new line.
        Shape textBox2 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        textBox2.TextBox.FitShapeToText = false;

        // Link the first text box to the second.
        textBox1.TextBox.Next = textBox2.TextBox;

        // Write the long text into the first text box.
        builder.MoveTo(textBox1.FirstParagraph);
        builder.Write(longText);

        // Save the document.
        doc.Save("LinkedTextBoxes.docx");
    }
}

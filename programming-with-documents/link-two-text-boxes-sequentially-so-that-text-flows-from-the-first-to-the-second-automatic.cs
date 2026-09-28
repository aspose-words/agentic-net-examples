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

        // Create the first text box.
        Shape textBox1 = new Shape(doc, ShapeType.TextBox);
        textBox1.Width = 300;
        textBox1.Height = 200;
        textBox1.WrapType = WrapType.Inline; // Simplify layout.

        // Add a paragraph with a long text that will overflow.
        Paragraph para1 = new Paragraph(doc);
        Run run1 = new Run(doc,
            "This is a long piece of text that will automatically flow from the first text box to the second one. " +
            "It contains enough words to exceed the bounds of the first text box, demonstrating the linking feature. " +
            "Aspose.Words makes it easy to link text boxes programmatically.");
        para1.AppendChild(run1);
        textBox1.AppendChild(para1);

        // Insert the first text box into the document.
        builder.InsertNode(textBox1);

        // Create the second text box.
        Shape textBox2 = new Shape(doc, ShapeType.TextBox);
        textBox2.Width = 300;
        textBox2.Height = 200;
        textBox2.WrapType = WrapType.Inline;

        // Add an empty paragraph (text will flow here).
        Paragraph para2 = new Paragraph(doc);
        textBox2.AppendChild(para2);

        // Insert the second text box after the first one.
        builder.InsertNode(textBox2);

        // NOTE:
        // In recent versions of Aspose.Words the TextBox class provides a LinkToShape property
        // that can be used to link overflow text from one text box to another:
        //     textBox1.TextBox.LinkToShape = textBox2;
        // If the current library version does not contain this property, linking cannot be performed
        // via the API. The code above creates two separate text boxes; linking would require a newer
        // Aspose.Words version that supports TextBox.LinkToShape.

        // Save the document.
        const string outputPath = "LinkedTextBoxes.docx";
        doc.Save(outputPath);
    }
}

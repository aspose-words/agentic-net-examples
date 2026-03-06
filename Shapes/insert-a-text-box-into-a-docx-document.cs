using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating text box shape with the desired size.
        Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 150, 100);
        // Example: make the text box not wrap with surrounding text.
        textBoxShape.WrapType = WrapType.None;

        // Move the builder's cursor inside the text box and add some text.
        builder.MoveTo(textBoxShape.LastParagraph);
        builder.Writeln("Hello world!");
        builder.Write("Hello again!");

        // Save the document to a DOCX file.
        doc.Save("TextBoxExample.docx");
    }
}

using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("Input.docx");

        // Create a DocumentBuilder to edit the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field at the current cursor position.
        // Parameters: field name, field type, format string, default display text, max length (0 = unlimited).
        builder.InsertTextInput("MyTextField", TextFormFieldType.Regular, "", "Enter your name", 0);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}

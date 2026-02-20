using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document("input.docx");

        // Create a DocumentBuilder attached to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the end of the document (or any desired location).
        builder.MoveToDocumentEnd();

        // Insert a text input form field.
        // Parameters: field name, field type, text format (empty for default), default text, maximum length.
        builder.InsertTextInput("MyTextInput", TextFormFieldType.Regular, "", "Enter your name", 30);

        // Save the modified document.
        doc.Save("output.docx");
    }
}

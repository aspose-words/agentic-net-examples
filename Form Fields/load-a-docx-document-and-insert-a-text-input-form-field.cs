using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the existing DOCX file.
        Document doc = new Document("input.docx");

        // Attach a DocumentBuilder to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a regular text input form field.
        // Parameters: field name, field type, text format (empty), default text, maximum length.
        builder.InsertTextInput("MyTextInput", TextFormFieldType.Regular, "", "Enter your name", 30);

        // Save the document with the new form field.
        doc.Save("output.docx");
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Fields;

class AddTextInputFormField
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("InputDocument.docx");

        // Create a DocumentBuilder to edit the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field at the current cursor position.
        // Parameters: name, type, format, default text, max length (0 = unlimited).
        builder.InsertTextInput("MyTextInput", TextFormFieldType.Regular, "", "Enter your text here", 0);

        // Save the modified document.
        doc.Save("OutputDocument.docx");
    }
}

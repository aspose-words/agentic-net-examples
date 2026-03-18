using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document – the builder provides the cursor for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field at the current cursor position.
        // Parameters:
        //   name          – logical name of the form field (also creates a bookmark with the same name).
        //   type          – the kind of text field; here we use a regular text field.
        //   format        – format string for the field value (empty for plain text).
        //   fieldValue    – placeholder text that appears until the user types something.
        //   maxLength     – maximum characters allowed (0 means unlimited).
        builder.InsertTextInput("TextInput", TextFormFieldType.Regular, "", "Enter your text here", 0);

        // Save the document to a file.
        doc.Save("TextInputFormField.docx");
    }
}

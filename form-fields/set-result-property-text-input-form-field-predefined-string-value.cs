using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field at the current cursor position.
        // Parameters: name, type, format, initial text, max length (0 = unlimited).
        FormField textField = builder.InsertTextInput(
            "MyTextInput",                     // field name
            TextFormFieldType.Regular,         // regular text field
            "",                                // no specific format
            "",                                // empty placeholder
            0);                                // unlimited length

        // Set the result (the value that the field holds) to a predefined string.
        textField.Result = "Predefined value";

        // Save the document to disk.
        doc.Save("ResultFormField.docx");
    }
}

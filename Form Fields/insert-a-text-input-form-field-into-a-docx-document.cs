using System;
using Aspose.Words;
using Aspose.Words.Fields;

class InsertTextInputFormField
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a prompt before the form field.
        builder.Write("Please enter your name: ");

        // Insert a text input form field at the current cursor position.
        // Parameters:
        //   name          – the form field name (also creates a bookmark with the same name).
        //   type          – the type of the text field (Regular allows any text).
        //   format        – format string for the field value (empty for no special formatting).
        //   fieldValue    – placeholder text shown inside the field.
        //   maxLength     – maximum number of characters (0 means unlimited).
        builder.InsertTextInput(
            name: "UserName",
            type: TextFormFieldType.Regular,
            format: "",
            fieldValue: "Enter your name here",
            maxLength: 0);

        // Save the document to a DOCX file.
        doc.Save("TextInputFormField.docx");
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Fields;

class InsertTextInputFormField
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a prompt before the form field.
        builder.Write("Please enter your name: ");

        // Insert a text input form field.
        // Parameters:
        //   name          – the form field name (also creates a bookmark with the same name).
        //   type          – the type of the text form field (regular text in this case).
        //   format        – optional format string (empty for no special formatting).
        //   fieldValue    – placeholder text shown to the user.
        //   maxLength     – maximum number of characters (0 = unlimited).
        builder.InsertTextInput(
            name: "NameField",
            type: TextFormFieldType.Regular,
            format: "",
            fieldValue: "Enter your name here",
            maxLength: 0);

        // Save the document to a DOCX file.
        doc.Save("TextInputFormField.docx");
    }
}

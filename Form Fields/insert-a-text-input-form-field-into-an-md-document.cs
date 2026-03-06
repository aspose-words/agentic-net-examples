using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a prompt before the form field.
        builder.Write("Please enter your name: ");

        // Insert a text input form field.
        // Parameters:
        //   name          – bookmark name (optional, can be empty).
        //   type          – type of the text field (regular text in this case).
        //   format        – format string (empty for no special formatting).
        //   fieldValue    – placeholder text shown to the user.
        //   maxLength     – 0 means unlimited length.
        builder.InsertTextInput(
            name: "UserName",
            type: TextFormFieldType.Regular,
            format: "",
            fieldValue: "Enter your name here",
            maxLength: 0);

        // Save the document to a .docx file.
        doc.Save("TextInputFormField.docx");
    }
}

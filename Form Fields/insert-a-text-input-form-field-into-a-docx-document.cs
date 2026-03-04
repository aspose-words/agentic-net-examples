using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a prompt before the form field.
        builder.Write("Please enter your name: ");

        // Insert a regular text input form field.
        // Parameters:
        //   name          – identifier for the field (also creates a bookmark with the same name).
        //   type          – type of the text field (Regular allows any text).
        //   format        – format string for the field value (empty for no special formatting).
        //   fieldValue    – placeholder text shown to the user.
        //   maxLength     – maximum characters allowed (0 means unlimited).
        builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "Enter name here", 0);

        // Save the document as a DOCX file.
        doc.Save("TextInputFormField.docx");
    }
}

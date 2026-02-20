using System;
using Aspose.Words;
using Aspose.Words.Fields;

class InsertTextInputFormField
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a regular text input form field.
        // Parameters:
        //   name               – The name of the form field.
        //   type               – The type of the text form field (Regular allows any text).
        //   format             – The text format string (empty for default).
        //   defaultText        – Prompt text shown when the field is empty.
        //   maxLength          – Maximum number of characters (0 means no limit, here 30).
        builder.InsertTextInput("MyTextInput", TextFormFieldType.Regular, "", "Enter your name here", 30);

        // Save the document to a DOCX file.
        doc.Save("TextInputFormField.docx");
    }
}

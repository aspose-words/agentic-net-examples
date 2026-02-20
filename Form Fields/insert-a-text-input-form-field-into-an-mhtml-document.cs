using System;
using Aspose.Words;
using Aspose.Words.Fields;

class InsertTextInputFormField
{
    static void Main()
    {
        // Load the existing MHTML document.
        Document doc = new Document("InputDocument.mhtml");

        // Create a DocumentBuilder to modify the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph to hold the form field (optional).
        builder.Writeln("Please enter your name:");

        // Insert a text input form field.
        // Parameters: name, type, format, default text, maximum length.
        builder.InsertTextInput(
            "UserName",                     // field name
            TextFormFieldType.Regular,      // allow any text
            "",                             // text format (none)
            "Enter name here",              // placeholder / default text
            50);                            // maximum length

        // Save the modified document back to MHTML format.
        doc.Save("OutputDocument.mhtml");
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a prompt before the form field.
        builder.Writeln("Enter your name:");

        // Insert a text input form field.
        // Parameters: name, type, default text, placeholder text, maximum length.
        FormField textInput = builder.InsertTextInput(
            "UserName",                     // field name
            TextFormFieldType.Regular,      // field type (any text)
            "",                             // default text (empty)
            "Enter name here",              // placeholder text shown to the user
            50);                            // maximum length of the input

        // Optional: set additional properties.
        textInput.TextInputFormat = "FIRST CAPITAL"; // display format
        textInput.MaxLength = 50;                    // enforce length

        // Save the document as a macro‑enabled DOCM file.
        doc.Save("TextInputFormField.docm");
    }
}

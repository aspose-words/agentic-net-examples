using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace InsertTextInputFormField
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new Word document (DOCM format)
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a text input form field at the current cursor position
            // Parameters: name, type, format, default text, max length (0 = unlimited)
            FormField textField = builder.InsertTextInput(
                "UserInput",                 // name of the form field (creates a bookmark with the same name)
                TextFormFieldType.Regular,   // field type: regular text
                "",                          // format string (empty for no special formatting)
                "Enter your text here",    // placeholder text shown to the user
                0);                           // maxLength = 0 means unlimited length

            // Optionally, set additional properties on the form field
            textField.TextInputDefault = "Enter your text here"; // default value when the field is empty
            textField.TextInputFormat = "";                     // no special text formatting

            // Save the document as a macro‑enabled DOCM file
            doc.Save("FormWithTextInput.docm", SaveFormat.Docm);
        }
    }
}

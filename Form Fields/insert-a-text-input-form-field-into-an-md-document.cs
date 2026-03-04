using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty Word document.
        Document doc = new Document();

        // Initialize a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a prompt before the form field.
        builder.Write("Please enter your name: ");

        // Insert a text input form field at the current cursor position.
        // Parameters:
        //   name          – bookmark name for the field (can be empty, but we give a name).
        //   type          – type of the text field (Regular allows any text).
        //   format        – format string (empty for no special formatting).
        //   fieldValue    – placeholder text shown to the user.
        //   maxLength     – 0 means unlimited length.
        builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "Enter name here", 0);

        // Save the document as a Markdown file.
        doc.Save("FormField.md");
    }
}

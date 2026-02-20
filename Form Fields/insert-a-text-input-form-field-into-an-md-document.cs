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

        // Add a line of text that will appear before the form field.
        builder.Writeln("Please enter your name:");

        // Insert a text input form field.
        // Parameters:
        //   name            – internal name of the field.
        //   type            – type of the text field (Regular allows any text).
        //   format          – optional format string (empty for none).
        //   defaultText     – placeholder text shown to the user.
        //   maxLength       – maximum number of characters (0 = unlimited, 30 here).
        builder.InsertTextInput("UserName", TextFormFieldType.Regular, "", "Enter name here", 30);

        // Save the document in Markdown format.
        doc.Save("FormFieldDocument.md", SaveFormat.Markdown);
    }
}

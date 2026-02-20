using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document("Input.docx");

        // Create a DocumentBuilder attached to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a prompt before the form field.
        builder.Writeln("Please enter your name:");

        // Insert a text input form field.
        // Parameters:
        //   name          – field name,
        //   type          – TextFormFieldType (Regular allows any text),
        //   format        – custom format string (empty for none),
        //   defaultText   – placeholder text shown to the user,
        //   maxLength     – maximum number of characters (0 = unlimited, here 50).
        builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "Enter name", 50);

        // Save the document with the new form field.
        doc.Save("Output.docx");
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("input.docx");

        // Create a DocumentBuilder for editing the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field at the current cursor position.
        // Parameters: name, type, format string, default display text, max length (0 = unlimited).
        builder.InsertTextInput("MyTextField", TextFormFieldType.Regular, "", "Enter text here", 0);

        // Save the modified document.
        doc.Save("output.docx");
    }
}

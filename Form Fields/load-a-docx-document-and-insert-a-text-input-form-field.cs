using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document("input.docx");

        // Create a DocumentBuilder to modify the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Position the builder at the end of the document (or any other location as needed).
        builder.MoveToDocumentEnd();

        // Insert a text input form field.
        // Parameters: name, type, format, default text, max length (0 = unlimited).
        builder.InsertTextInput("MyTextInput", TextFormFieldType.Regular, "", "Enter text here", 0);

        // Save the modified document.
        doc.Save("output.docx");
    }
}

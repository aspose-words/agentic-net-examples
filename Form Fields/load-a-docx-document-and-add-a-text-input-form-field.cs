using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document("InputDocument.docx");

        // Create a DocumentBuilder for the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the end of the document (or any desired position).
        builder.MoveToDocumentEnd();

        // Insert a text input form field.
        // Parameters: name, type, format, default text, max length (0 = unlimited).
        builder.InsertTextInput(
            name: "MyTextInput",
            type: TextFormFieldType.Regular,
            format: "",
            fieldValue: "Enter your text here",
            maxLength: 0);

        // Save the modified document.
        doc.Save("OutputDocument.docx");
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("input.docx");

        // Create a DocumentBuilder to edit the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the end of the document (or any desired position).
        builder.MoveToDocumentEnd();

        // Insert a checkbox form field.
        // Parameters: name, default checked value, size (0 = auto size).
        builder.InsertCheckBox("MyCheckBox", false, 0);

        // Save the modified document.
        doc.Save("output.docx");
    }
}

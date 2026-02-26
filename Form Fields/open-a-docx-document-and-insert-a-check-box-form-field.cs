using System;
using Aspose.Words;

class InsertCheckBoxExample
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document("InputDocument.docx");

        // Create a DocumentBuilder to work with the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the end of the document (or any desired position).
        builder.MoveToDocumentEnd();

        // Insert a checkbox form field.
        // Parameters: name, defaultValue, checkedValue, size (0 = auto size).
        // Here we give the field a name, set its default unchecked, current unchecked, and let Word calculate size.
        builder.InsertCheckBox("MyCheckBox", false, false, 0);

        // Save the modified document.
        doc.Save("OutputDocument.docx");
    }
}

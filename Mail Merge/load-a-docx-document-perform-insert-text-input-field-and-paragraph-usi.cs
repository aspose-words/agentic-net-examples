using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document("input.docx");

        // Attach a DocumentBuilder to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a regular paragraph of text.
        builder.Writeln("This paragraph is added by DocumentBuilder.");

        // Insert a text input form field.
        // Parameters: name, field type, format (empty), default text, max length (0 = unlimited).
        builder.InsertTextInput("MyTextField", TextFormFieldType.Regular, "", "Default value", 0);

        // Insert a new paragraph break after the form field.
        builder.Writeln();

        // Insert a PAGE field and update it immediately.
        builder.InsertField(FieldType.FieldPage, true);

        // Add another paragraph to finish the inserted content.
        builder.Writeln("End of inserted content.");

        // Save the document as a PNG image (first page rendered as PNG).
        doc.Save("output.png", SaveFormat.Png);
    }
}

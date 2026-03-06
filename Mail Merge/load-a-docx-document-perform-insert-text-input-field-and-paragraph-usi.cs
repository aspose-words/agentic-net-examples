using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document("Input.docx");

        // Attach a DocumentBuilder to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field.
        // Parameters: name, field type, default text, format string, max length (0 = unlimited).
        builder.InsertTextInput("MyTextField", TextFormFieldType.Regular, "Default text", "", 0);

        // Insert a DATE field that displays the current date in a specific format.
        builder.InsertField("DATE \\@ \"MMMM d, yyyy\"");

        // Insert a new paragraph break.
        builder.InsertParagraph();

        // Write a line of text into the newly created paragraph.
        builder.Writeln("This is a new paragraph added after the field.");

        // Save the document as a PNG image (renders the first page).
        doc.Save("Output.png", SaveFormat.Png);
    }
}

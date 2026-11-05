using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field.
        // Parameters: name, field type, format, default text (used here as placeholder), max length (0 = no limit).
        builder.InsertTextInput(
            "TextInputField",
            TextFormFieldType.Regular,
            "",
            "Enter your name",
            0);

        // Save the document to disk.
        doc.Save("FormFieldExample.docx");
    }
}

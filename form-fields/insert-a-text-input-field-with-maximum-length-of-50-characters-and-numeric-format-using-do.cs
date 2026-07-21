using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a prompt for the user.
        builder.Writeln("Enter a numeric value (max 50 characters):");

        // Insert a numeric text input form field.
        // Parameters: name, type (Number), format (empty), default value, maximum length.
        FormField numericField = builder.InsertTextInput(
            "NumericInput",
            TextFormFieldType.Number,
            "",
            "0",
            50);

        // Ensure the field is enabled.
        numericField.Enabled = true;

        // Save the document to disk.
        doc.Save("FormFieldNumeric.docx");
    }
}

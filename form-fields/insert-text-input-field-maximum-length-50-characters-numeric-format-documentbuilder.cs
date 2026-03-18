using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a numeric text input form field.
        // Parameters:
        //   name       – field name (also creates a bookmark with the same name).
        //   type       – TextFormFieldType.Number to restrict input to numbers.
        //   format     – numeric format string; "0" means integer without extra formatting.
        //   fieldValue – default placeholder text (empty in this case).
        //   maxLength  – maximum number of characters the user can type (50).
        FormField numericField = builder.InsertTextInput(
            "NumericInput",
            TextFormFieldType.Number,
            "0",
            "",
            50);

        // Optionally, you can set additional properties on the field.
        // For example, ensure the field is enabled.
        numericField.Enabled = true;

        // Save the document to a file.
        doc.Save("TextInputNumeric.docx");
    }
}

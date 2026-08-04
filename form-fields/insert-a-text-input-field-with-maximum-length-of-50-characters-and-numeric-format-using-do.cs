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

        // Insert a numeric text input form field.
        // Parameters: name, type (Number), format (empty for default), default value, max length.
        FormField numericField = builder.InsertTextInput(
            name: "NumericInput",
            type: TextFormFieldType.Number,
            format: "",
            fieldValue: "0",
            maxLength: 50);

        // Optionally set a display format (e.g., integer without decimals).
        numericField.TextInputFormat = "0";

        // Save the document to a file.
        doc.Save("FormField.docx");
    }
}

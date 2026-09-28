using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a numeric text input field with a maximum length of 50 characters.
        // The builder's InsertTextInput method already applies the max length.
        string fieldName = "NumericInput";
        builder.InsertTextInput(fieldName, TextFormFieldType.Number, "", "0", 50);

        // Save the document to verify the result.
        string outputPath = "FormFieldExample.docx";
        doc.Save(outputPath);

        // Validate that the field exists and has the correct settings.
        FormField field = doc.Range.FormFields[fieldName];
        if (field == null)
        {
            throw new InvalidOperationException($"Form field '{fieldName}' was not found.");
        }

        // Verify the field is configured for numeric input.
        if (field.TextInputType != TextFormFieldType.Number)
        {
            throw new InvalidOperationException("The field is not configured for numeric input.");
        }

        // Verify the default value (Result) is set to "0".
        if (field.Result != "0")
        {
            throw new InvalidOperationException("The field's default value is not set to '0'.");
        }

        Console.WriteLine("Form field created and validated successfully.");
    }
}

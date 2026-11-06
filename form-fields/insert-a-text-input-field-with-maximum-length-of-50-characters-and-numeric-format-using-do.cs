using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a numeric text input form field with a maximum length of 50 characters.
        // Parameters: name, type, format, default text, max length.
        builder.InsertTextInput("NumericField", TextFormFieldType.Number, "", "0", 50);

        // Save the initial document.
        const string fileName = "FormFieldExample.docx";
        doc.Save(fileName);

        // Retrieve the form field by name.
        FormField? field = doc.Range.FormFields["NumericField"];
        if (field == null)
        {
            throw new InvalidOperationException("The form field 'NumericField' was not found.");
        }

        // Validate the default value.
        if (field.Result != "0")
        {
            throw new InvalidOperationException($"Unexpected default value: {field.Result}");
        }

        // Update the field's result.
        field.Result = "12345";

        // Validate the updated value.
        if (field.Result != "12345")
        {
            throw new InvalidOperationException($"Failed to update the field value: {field.Result}");
        }

        // Save the document after the update.
        doc.Save(fileName);
    }
}

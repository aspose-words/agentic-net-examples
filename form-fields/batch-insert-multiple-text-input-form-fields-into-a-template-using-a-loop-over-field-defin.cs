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

        // Define the text input form fields to be inserted.
        var fieldDefinitions = new[]
        {
            new { Name = "FirstName", DefaultValue = "John", MaxLength = 20 },
            new { Name = "LastName",  DefaultValue = "Doe",  MaxLength = 20 },
            new { Name = "Email",     DefaultValue = "example@example.com", MaxLength = 50 },
            new { Name = "Phone",     DefaultValue = "",    MaxLength = 15 }
        };

        // Insert each field into the document.
        foreach (var def in fieldDefinitions)
        {
            // Add a label for the field.
            builder.Writeln($"Enter {def.Name}:");

            // Insert the text input form field.
            FormField field = builder.InsertTextInput(
                def.Name,                     // field name
                TextFormFieldType.Regular,    // field type
                "",                           // format (none)
                def.DefaultValue,             // initial text
                def.MaxLength);               // maximum length

            // Ensure the field's result matches the default value.
            field.Result = def.DefaultValue;
        }

        // Validate that at least one form field exists.
        FormFieldCollection formFields = doc.Range.FormFields;
        if (formFields.Count == 0)
            throw new InvalidOperationException("No form fields were created.");

        // Verify each defined field exists and contains the expected value.
        foreach (var def in fieldDefinitions)
        {
            FormField field = formFields[def.Name];
            if (field == null)
                throw new InvalidOperationException($"Form field '{def.Name}' not found.");

            if (field.Result != def.DefaultValue)
                throw new InvalidOperationException($"Form field '{def.Name}' result mismatch.");
        }

        // Save the document to disk.
        doc.Save("FormFieldsBatch.docx");
    }
}

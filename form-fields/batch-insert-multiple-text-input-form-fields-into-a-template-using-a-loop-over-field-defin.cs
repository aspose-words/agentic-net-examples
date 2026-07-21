using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Output file name.
        const string outputPath = "BatchFormFields.docx";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define the form fields to be inserted.
        var fieldDefinitions = new[]
        {
            new { Name = "FirstName", Placeholder = "Enter first name", MaxLength = 30 },
            new { Name = "LastName",  Placeholder = "Enter last name",  MaxLength = 30 },
            new { Name = "Email",     Placeholder = "Enter email address", MaxLength = 50 },
            new { Name = "Phone",     Placeholder = "Enter phone number", MaxLength = 20 }
        };

        // Introductory text.
        builder.Writeln("Please fill out the following form:");
        builder.Writeln();

        // Insert each text input form field using a loop.
        foreach (var def in fieldDefinitions)
        {
            // Write a label for the field.
            builder.Write($"{def.Name}: ");

            // Insert the text input form field.
            FormField field = builder.InsertTextInput(
                def.Name,                     // field name
                TextFormFieldType.Regular,    // field type
                "",                           // format (none)
                def.Placeholder,              // placeholder text shown to the user
                def.MaxLength);               // maximum length (0 = unlimited)

            // Ensure the field has an explicit default value (empty string).
            field.TextInputDefault = string.Empty;

            // Move to the next line after the field.
            builder.Writeln();
        }

        // Validate that at least one form field exists.
        FormFieldCollection formFields = doc.Range.FormFields;
        if (formFields.Count == 0)
            throw new InvalidOperationException("No form fields were inserted into the document.");

        // Verify each defined field exists and optionally set a sample result.
        foreach (var def in fieldDefinitions)
        {
            FormField f = formFields[def.Name];
            if (f == null)
                throw new InvalidOperationException($"Form field '{def.Name}' was not found in the document.");

            // Example: set a sample value for demonstration purposes.
            f.Result = $"Sample {def.Name}";
        }

        // Save the document to disk.
        doc.Save(outputPath);
    }
}

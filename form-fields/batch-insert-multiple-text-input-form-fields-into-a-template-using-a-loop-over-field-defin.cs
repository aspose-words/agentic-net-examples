using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

public class BatchFormFieldsExample
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define the form fields to be inserted.
        var fieldDefinitions = new List<FieldDefinition>
        {
            new FieldDefinition { Name = "FirstName", Placeholder = "Enter first name", MaxLength = 30 },
            new FieldDefinition { Name = "LastName", Placeholder = "Enter last name", MaxLength = 30 },
            new FieldDefinition { Name = "Email", Placeholder = "Enter email address", MaxLength = 50 },
            new FieldDefinition { Name = "Phone", Placeholder = "Enter phone number", MaxLength = 20 }
        };

        // Insert each text input form field using a loop.
        foreach (var def in fieldDefinitions)
        {
            // Write a prompt for the user.
            builder.Writeln($"Please provide {def.Name}:");

            // Insert the text input form field.
            FormField field = builder.InsertTextInput(
                def.Name,                     // field name (required)
                TextFormFieldType.Regular,   // type of the text field
                "",                          // format string (none)
                def.Placeholder,             // initial placeholder text
                def.MaxLength);              // maximum length (0 = unlimited)

            // Optionally set a default value (same as placeholder in this case).
            field.TextInputDefault = def.Placeholder;
        }

        // Validate that at least one form field was created.
        FormFieldCollection formFields = doc.Range.FormFields;
        if (formFields.Count == 0)
            throw new InvalidOperationException("No form fields were inserted into the document.");

        // Update each field with a sample value and verify the assignment.
        foreach (var def in fieldDefinitions)
        {
            FormField field = formFields[def.Name];
            if (field == null)
                throw new InvalidOperationException($"Form field '{def.Name}' was not found.");

            // Assign a new value to the field.
            string newValue = $"Sample {def.Name}";
            field.Result = newValue;

            // Validate that the value was set correctly.
            if (field.Result != newValue)
                throw new InvalidOperationException($"Failed to set value for field '{def.Name}'.");
        }

        // Save the document to disk.
        doc.Save("BatchFormFields.docx");
    }

    // Simple class to hold field definition data.
    private class FieldDefinition
    {
        public string Name { get; set; }
        public string Placeholder { get; set; }
        public int MaxLength { get; set; }
    }
}

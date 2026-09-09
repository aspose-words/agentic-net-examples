using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    // Simple DTO to hold field definition data.
    private class FieldDefinition
    {
        public string Name { get; set; }          // Form field name (also bookmark name).
        public string Placeholder { get; set; }   // Text shown when the field is empty.
        public int MaxLength { get; set; }        // Maximum characters allowed (0 = unlimited).
    }

    public static void Main()
    {
        // Prepare a list of field definitions to be inserted.
        var fields = new List<FieldDefinition>
        {
            new FieldDefinition { Name = "FirstName", Placeholder = "Enter first name", MaxLength = 30 },
            new FieldDefinition { Name = "LastName", Placeholder = "Enter last name", MaxLength = 30 },
            new FieldDefinition { Name = "Email", Placeholder = "example@domain.com", MaxLength = 50 },
            new FieldDefinition { Name = "Phone", Placeholder = "123-456-7890", MaxLength = 20 }
        };

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each text input form field using the definitions above.
        foreach (var def in fields)
        {
            // Write a label for the field.
            builder.Writeln($"{def.Name}:");

            // Insert the text input form field.
            // Parameters: name, type, format (empty), default text, max length.
            builder.InsertTextInput(def.Name, TextFormFieldType.Regular, "", def.Placeholder, def.MaxLength);
        }

        // Validate that form fields were added.
        FormFieldCollection formFields = doc.Range.FormFields;
        if (formFields.Count == 0)
            throw new InvalidOperationException("No form fields were inserted into the document.");

        // Optionally, verify each field exists by name.
        foreach (var def in fields)
        {
            FormField? field = formFields[def.Name];
            if (field == null)
                throw new InvalidOperationException($"Form field '{def.Name}' was not found.");
        }

        // Save the document to disk.
        const string outputPath = "BatchFormFields.docx";
        doc.Save(outputPath);
    }
}

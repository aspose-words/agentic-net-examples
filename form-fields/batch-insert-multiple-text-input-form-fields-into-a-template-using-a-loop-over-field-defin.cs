using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace BatchFormFieldsExample
{
    public class Program
    {
        // Definition of a text input form field.
        private class TextFieldDefinition
        {
            public string Name { get; }
            public string Placeholder { get; }
            public int MaxLength { get; }

            public TextFieldDefinition(string name, string placeholder, int maxLength)
            {
                Name = name;
                Placeholder = placeholder;
                MaxLength = maxLength;
            }
        }

        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Define the fields to be inserted.
            TextFieldDefinition[] fields = new[]
            {
                new TextFieldDefinition("FirstName", "Enter first name", 50),
                new TextFieldDefinition("LastName", "Enter last name", 50),
                new TextFieldDefinition("Email", "Enter email address", 100),
                new TextFieldDefinition("Phone", "Enter phone number", 20)
            };

            // Insert each field into the document.
            foreach (var fieldDef in fields)
            {
                // Write a label for the field.
                builder.Writeln($"{fieldDef.Name}:");

                // Insert the text input form field.
                FormField formField = builder.InsertTextInput(
                    fieldDef.Name,                     // field name
                    TextFormFieldType.Regular,         // field type
                    "",                                // format (none)
                    fieldDef.Placeholder,              // default visible text
                    fieldDef.MaxLength);               // maximum length

                // Optional: set a default value programmatically.
                formField.SetTextInputValue(string.Empty);
            }

            // Validate that all fields were added.
            FormFieldCollection formFields = doc.Range.FormFields;
            if (formFields.Count != fields.Length)
                throw new InvalidOperationException("Not all form fields were inserted.");

            // Save the document.
            doc.Save("BatchFormFields.docx");
        }
    }
}

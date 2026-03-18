using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

namespace FormFieldBatchInsert
{
    // Represents the definition of a text input form field.
    public class FieldDefinition
    {
        public string Name { get; set; }                 // Bookmark/form field name.
        public TextFormFieldType Type { get; set; }      // Regular, Number, Date, etc.
        public string Format { get; set; }               // Formatting string (e.g., "UPPERCASE").
        public string DefaultValue { get; set; }         // Placeholder or default text.
        public int MaxLength { get; set; }               // 0 for unlimited length.
    }

    public static class FormFieldInserter
    {
        // Inserts a batch of text input form fields into a template document.
        public static void InsertFormFields(string templatePath, string outputPath, List<FieldDefinition> fieldDefs)
        {
            // Load the existing template.
            Document doc = new Document(templatePath);

            // Create a DocumentBuilder positioned at the end of the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();

            // Insert each form field defined in the list.
            foreach (var def in fieldDefs)
            {
                // Write a label before the field for readability (optional).
                builder.Writeln($"Please enter {def.Name}:");

                // Insert the text input form field using the Aspose.Words API.
                // Parameters: name, type, format, default value, max length.
                builder.InsertTextInput(def.Name, def.Type, def.Format ?? string.Empty,
                                        def.DefaultValue ?? string.Empty, def.MaxLength);

                // Add a paragraph break after each field.
                builder.Writeln();
            }

            // Save the modified document.
            doc.Save(outputPath);
        }

        // Example usage.
        public static void Main()
        {
            // Path to the template document (must exist).
            string templatePath = @"C:\Templates\MyFormTemplate.docx";

            // Desired output path.
            string outputPath = @"C:\Output\MyFormFilled.docx";

            // Define a collection of form fields to insert.
            var fields = new List<FieldDefinition>
            {
                new FieldDefinition
                {
                    Name = "FirstName",
                    Type = TextFormFieldType.Regular,
                    Format = "FIRST CAPITAL",
                    DefaultValue = "Enter first name",
                    MaxLength = 0
                },
                new FieldDefinition
                {
                    Name = "LastName",
                    Type = TextFormFieldType.Regular,
                    Format = "FIRST CAPITAL",
                    DefaultValue = "Enter last name",
                    MaxLength = 0
                },
                new FieldDefinition
                {
                    Name = "Age",
                    Type = TextFormFieldType.Number,
                    Format = "",
                    DefaultValue = "0",
                    MaxLength = 3
                },
                new FieldDefinition
                {
                    Name = "BirthDate",
                    Type = TextFormFieldType.Date,
                    Format = "MM/dd/yyyy",
                    DefaultValue = "",
                    MaxLength = 0
                }
            };

            // Perform the batch insertion.
            InsertFormFields(templatePath, outputPath, fields);
        }
    }
}

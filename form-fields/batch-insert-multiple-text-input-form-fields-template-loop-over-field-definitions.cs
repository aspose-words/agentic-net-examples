using System;
using System.Collections.Generic;
using System.IO;
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
            Document doc;

            // Load the existing template if it exists; otherwise create a new blank document.
            if (File.Exists(templatePath))
            {
                doc = new Document(templatePath);
            }
            else
            {
                doc = new Document();
                // Add a single empty paragraph so the builder has a place to start.
                doc.FirstSection.Body.AppendChild(new Paragraph(doc));
            }

            // Ensure the output directory exists.
            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

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
            // Use paths relative to the current working directory.
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "MyFormTemplate.docx");
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MyFormFilled.docx");

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

            Console.WriteLine($"Form fields inserted. Output saved to: {outputPath}");
        }
    }
}

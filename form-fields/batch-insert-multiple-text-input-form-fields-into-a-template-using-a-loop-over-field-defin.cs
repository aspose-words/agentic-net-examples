using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

namespace FormFieldsBatchInsert
{
    public class Program
    {
        public static void Main()
        {
            // Define the form field specifications.
            var fieldDefinitions = new List<(string Name, string DefaultValue)>
            {
                ("FirstName", "John"),
                ("LastName", "Doe"),
                ("Email", "example@example.com")
            };

            // Create a new blank document and a DocumentBuilder.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Insert a heading.
            builder.Writeln("User Information Form");
            builder.Writeln();

            // Batch insert text input form fields using the definitions.
            foreach (var (name, defaultValue) in fieldDefinitions)
            {
                // Insert a label for the field.
                builder.Write($"{name}: ");

                // Insert the text input form field.
                builder.InsertTextInput(name, TextFormFieldType.Regular, "", defaultValue, 0);

                // Move to the next line after each field.
                builder.Writeln();
            }

            // Save the document with the inserted form fields.
            const string outputPath = "BatchFormFields.docx";
            doc.Save(outputPath);

            // Load the saved document to verify the fields.
            var loadedDoc = new Document(outputPath);
            var formFields = loadedDoc.Range.FormFields;

            // Validate that each expected field exists and output its default value.
            foreach (var (name, _) in fieldDefinitions)
            {
                var field = formFields[name];
                if (field == null)
                {
                    throw new InvalidOperationException($"Form field '{name}' was not found in the document.");
                }

                // Output the field name and its current result (default value).
                Console.WriteLine($"{field.Name}: {field.Result}");
            }
        }
    }
}

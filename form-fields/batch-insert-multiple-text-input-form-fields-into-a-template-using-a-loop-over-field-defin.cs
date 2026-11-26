using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a set of text input form fields to insert.
        var fieldDefinitions = new List<(string Name, string Label, string DefaultText, int MaxLength)>
        {
            ("FirstName", "First Name:", "John", 50),
            ("LastName", "Last Name:", "Doe", 50),
            ("Email", "Email Address:", "john.doe@example.com", 100),
            ("Phone", "Phone Number:", "+1-555-1234", 20)
        };

        // Insert each form field using a loop.
        foreach (var def in fieldDefinitions)
        {
            // Write the label.
            builder.Writeln(def.Label);

            // Insert the text input form field.
            builder.InsertTextInput(
                def.Name,                     // field name
                TextFormFieldType.Regular,    // regular text field
                "",                           // format (empty for plain text)
                def.DefaultText,              // default text
                def.MaxLength);               // maximum length

            // Add a blank line after each field for readability.
            builder.Writeln();
        }

        // Save the document with the inserted form fields.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Validate that all expected form fields exist and have the correct default values.
        foreach (var def in fieldDefinitions)
        {
            FormField field = doc.Range.FormFields[def.Name];
            if (field == null)
                throw new InvalidOperationException($"Form field '{def.Name}' was not found.");

            if (field.Result != def.DefaultText)
                throw new InvalidOperationException($"Form field '{def.Name}' does not contain the expected default value.");
        }

        // Indicate successful completion (no interactive output required).
        Console.WriteLine($"Document '{outputPath}' created with {fieldDefinitions.Count} text input form fields.");
    }
}

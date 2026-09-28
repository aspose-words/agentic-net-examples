using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Define the path for the template and the final document.
        const string templatePath = "Template.docx";
        const string outputPath = "Result.docx";

        // Step 1: Create a simple template document if it does not exist.
        // The template contains a heading where form fields will be inserted.
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("Please fill in the following information:");
        templateBuilder.Writeln(); // Add an empty line for visual separation.
        templateDoc.Save(templatePath);

        // Step 2: Load the template document.
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the end of the document to start inserting fields.
        builder.MoveToDocumentEnd();

        // Step 3: Define a collection of text input form field specifications.
        var fieldDefinitions = new List<(string Name, string DefaultValue)>
        {
            ("FirstName", "John"),
            ("LastName", "Doe"),
            ("Email", "john.doe@example.com"),
            ("Phone", "123-456-7890")
        };

        // Step 4: Insert each text input form field using a loop.
        foreach (var (name, defaultValue) in fieldDefinitions)
        {
            // Add a label for the field.
            builder.Writeln($"{name}:");

            // Insert the text input form field.
            // TextFormFieldType.Regular creates a regular text input.
            builder.InsertTextInput(name, TextFormFieldType.Regular, "", defaultValue, 0);

            // Add a line break after each field for readability.
            builder.Writeln();
        }

        // Step 5: Validate that at least one form field was added.
        if (doc.Range.FormFields == null || doc.Range.FormFields.Count == 0)
        {
            throw new InvalidOperationException("No form fields were inserted into the document.");
        }

        // Optional: Verify that each inserted field contains the expected default value.
        foreach (var (name, defaultValue) in fieldDefinitions)
        {
            FormField? field = doc.Range.FormFields[name];
            if (field == null)
                throw new InvalidOperationException($"Form field '{name}' was not found.");

            if (field.Result != defaultValue)
                throw new InvalidOperationException($"Form field '{name}' does not contain the expected default value.");
        }

        // Step 6: Save the resulting document.
        doc.Save(outputPath);
    }
}

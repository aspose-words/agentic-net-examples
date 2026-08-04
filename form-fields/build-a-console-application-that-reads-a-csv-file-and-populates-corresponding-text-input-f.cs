using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Paths for the files.
        string csvPath = "data.csv";
        string templatePath = "template.docx";
        string outputPath = "filled.docx";

        // Create a simple CSV file with field names and values.
        // Format: FieldName,Value
        File.WriteAllLines(csvPath, new[]
        {
            "FirstName,John",
            "LastName,Doe",
            "Email,john.doe@example.com"
        });

        // Build a template document that contains text input form fields
        // matching the CSV field names.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Helper to insert a labeled text input field.
        void InsertLabeledField(string fieldName, string placeholder)
        {
            builder.Writeln($"{fieldName}:");
            // Insert a text input form field with a placeholder and no length limit.
            builder.InsertTextInput(fieldName, TextFormFieldType.Regular, "", placeholder, 0);
            builder.Writeln(); // Add a blank line after each field.
        }

        InsertLabeledField("FirstName", "Enter first name");
        InsertLabeledField("LastName", "Enter last name");
        InsertLabeledField("Email", "Enter email address");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template document.
        Document doc = new Document(templatePath);

        // Read CSV data into a dictionary.
        var data = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var line in File.ReadAllLines(csvPath))
        {
            if (string.IsNullOrWhiteSpace(line))
                continue;

            var parts = line.Split(new[] { ',' }, 2);
            if (parts.Length != 2)
                continue; // Skip malformed lines.

            string key = parts[0].Trim();
            string value = parts[1].Trim();
            data[key] = value;
        }

        // Ensure the document contains at least one form field.
        FormFieldCollection formFields = doc.Range.FormFields;
        if (formFields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        // Populate each form field with the corresponding CSV value.
        foreach (var kvp in data)
        {
            // Retrieve the form field by its bookmark/name.
            FormField field = formFields[kvp.Key];
            if (field == null)
                throw new KeyNotFoundException($"Form field '{kvp.Key}' not found in the document.");

            // For text input fields, use SetTextInputValue to apply formatting.
            field.SetTextInputValue(kvp.Value);

            // Validate that the value was set.
            if (!string.Equals(field.Result, kvp.Value, StringComparison.Ordinal))
                throw new InvalidOperationException($"Failed to set value for field '{kvp.Key}'.");
        }

        // Save the filled document.
        doc.Save(outputPath);
    }
}

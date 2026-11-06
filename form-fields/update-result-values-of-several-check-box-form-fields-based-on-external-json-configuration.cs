using System;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with several checkbox form fields.
        builder.Writeln("Please review the following options:");
        InsertCheckBox(builder, "OptionA", false);
        InsertCheckBox(builder, "OptionB", true);
        InsertCheckBox(builder, "OptionC", false);

        // Save the initial document (optional, can be omitted).
        doc.Save("FormFieldsTemplate.docx");

        // External JSON configuration that determines which checkboxes should be checked.
        string jsonConfig = @"{
            ""OptionA"": true,
            ""OptionB"": false,
            ""OptionC"": true
        }";

        // Parse the JSON into a dictionary.
        Dictionary<string, bool> config = JsonSerializer.Deserialize<Dictionary<string, bool>>(jsonConfig)
                                            ?? throw new InvalidOperationException("Failed to parse JSON configuration.");

        // Update checkbox form fields based on the configuration.
        foreach (KeyValuePair<string, bool> entry in config)
        {
            // Retrieve the form field by name.
            FormField field = doc.Range.FormFields[entry.Key];
            if (field == null)
                throw new InvalidOperationException($"Form field '{entry.Key}' was not found in the document.");

            // Ensure the field is a checkbox.
            if (field.Type != FieldType.FieldFormCheckBox)
                throw new InvalidOperationException($"Form field '{entry.Key}' is not a checkbox.");

            // Update the checked state.
            field.Checked = entry.Value;

            // Validate the update.
            if (field.Checked != entry.Value)
                throw new InvalidOperationException($"Failed to set the checked state for field '{entry.Key}'.");
        }

        // Save the updated document.
        doc.Save("FormFieldsUpdated.docx");
    }

    // Helper method to insert a checkbox with a given name and default state.
    private static void InsertCheckBox(DocumentBuilder builder, string name, bool isChecked)
    {
        // Insert the checkbox form field.
        builder.InsertCheckBox(name, isChecked, 20);
        // Add a space after the checkbox for readability.
        builder.Writeln($" {name}");
    }
}

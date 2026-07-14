using System;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // JSON configuration that maps form field names to their desired checked state.
        string jsonConfig = @"{ ""CheckBox1"": true, ""CheckBox2"": false, ""CheckBox3"": true }";

        // Deserialize the JSON into a dictionary for easy lookup.
        Dictionary<string, bool> config = JsonSerializer.Deserialize<Dictionary<string, bool>>(jsonConfig);

        // Create a new document and a builder to insert form fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three check box form fields with distinct names.
        builder.Write("Option 1: ");
        builder.InsertCheckBox("CheckBox1", false, 0);
        builder.Writeln();

        builder.Write("Option 2: ");
        builder.InsertCheckBox("CheckBox2", false, 0);
        builder.Writeln();

        builder.Write("Option 3: ");
        builder.InsertCheckBox("CheckBox3", false, 0);
        builder.Writeln();

        // Access the collection of form fields in the document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Ensure that the document contains at least one form field.
        if (formFields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        // Update each check box based on the external JSON configuration.
        foreach (FormField field in formFields)
        {
            // Process only check box fields.
            if (field.Type == FieldType.FieldFormCheckBox)
            {
                // Verify that the field has a name.
                if (string.IsNullOrEmpty(field.Name))
                    continue;

                // Retrieve the desired value from the configuration.
                if (!config.TryGetValue(field.Name, out bool desiredValue))
                    throw new KeyNotFoundException($"No configuration entry found for form field '{field.Name}'.");

                // Set the checked state.
                field.Checked = desiredValue;

                // Validate that the assignment succeeded.
                if (field.Checked != desiredValue)
                    throw new InvalidOperationException($"Failed to set the checked state for field '{field.Name}'.");
            }
        }

        // Update fields in the document (good practice before saving).
        doc.UpdateFields();

        // Save the updated document.
        const string outputPath = "UpdatedFormFields.docx";
        doc.Save(outputPath);
    }
}

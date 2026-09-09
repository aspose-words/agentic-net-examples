using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and add checkbox form fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three check boxes with distinct names.
        builder.Write("Option A: ");
        builder.InsertCheckBox("CheckBox1", false, 0);
        builder.InsertParagraph();

        builder.Write("Option B: ");
        builder.InsertCheckBox("CheckBox2", false, 0);
        builder.InsertParagraph();

        builder.Write("Option C: ");
        builder.InsertCheckBox("CheckBox3", false, 0);
        builder.InsertParagraph();

        // JSON configuration that maps field names to the desired checked state.
        string jsonConfig = @"{
            ""CheckBox1"": true,
            ""CheckBox2"": false,
            ""CheckBox3"": true
        }";

        // Parse the JSON into a dictionary.
        Dictionary<string, bool> config = JsonSerializer.Deserialize<Dictionary<string, bool>>(jsonConfig);

        // Update each form field according to the configuration.
        foreach (KeyValuePair<string, bool> kvp in config)
        {
            // Retrieve the form field by name.
            FormField field = doc.Range.FormFields[kvp.Key];
            if (field == null)
                throw new InvalidOperationException($"Form field '{kvp.Key}' not found.");

            // Ensure the field is a checkbox.
            if (field.Type != FieldType.FieldFormCheckBox)
                throw new InvalidOperationException($"Form field '{kvp.Key}' is not a checkbox.");

            // Set the checked state.
            field.Checked = kvp.Value;

            // Validate the assignment.
            if (field.Checked != kvp.Value)
                throw new InvalidOperationException($"Failed to set checked state for '{kvp.Key}'.");
        }

        // Save the updated document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "UpdatedFormFields.docx");
        doc.Save(outputPath);
    }
}

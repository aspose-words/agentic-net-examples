using System;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several checkbox form fields with default values.
        builder.Writeln("Please select the options below:");
        builder.InsertCheckBox("CheckBox1", false, 20);
        builder.Writeln(" Option 1");
        builder.InsertCheckBox("CheckBox2", true, 20);
        builder.Writeln(" Option 2");
        builder.InsertCheckBox("CheckBox3", false, 20);
        builder.Writeln(" Option 3");

        // Ensure at least one form field exists.
        if (doc.Range.FormFields.Count == 0)
            throw new InvalidOperationException("No form fields were created in the document.");

        // External JSON configuration (could be read from a file or service).
        string jsonConfig = @"{
            ""CheckBox1"": true,
            ""CheckBox2"": false,
            ""CheckBox3"": true
        }";

        // Parse JSON into a dictionary.
        Dictionary<string, bool> config = JsonSerializer.Deserialize<Dictionary<string, bool>>(jsonConfig)
                                            ?? throw new InvalidOperationException("Failed to parse JSON configuration.");

        // Update each checkbox based on the configuration.
        foreach (KeyValuePair<string, bool> kvp in config)
        {
            // Retrieve the form field by name.
            FormField field = doc.Range.FormFields[kvp.Key];
            if (field == null)
                throw new KeyNotFoundException($"Form field '{kvp.Key}' was not found in the document.");

            // Verify that the field is a checkbox.
            if (field.Type != FieldType.FieldFormCheckBox)
                throw new InvalidOperationException($"Form field '{kvp.Key}' is not a checkbox.");

            // Update the checked state.
            field.Checked = kvp.Value;

            // Validate the update.
            if (field.Checked != kvp.Value)
                throw new InvalidOperationException($"Failed to set checkbox '{kvp.Key}' to '{kvp.Value}'.");
        }

        // Save the updated document.
        const string outputPath = "UpdatedFormFields.docx";
        doc.Save(outputPath);
    }
}

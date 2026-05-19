using System;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // JSON configuration that maps form field names to the desired checked state.
        const string jsonConfig = @"{
            ""CheckBox1"": true,
            ""CheckBox2"": false,
            ""CheckBox3"": true
        }";

        // Parse the JSON into a dictionary for easy lookup.
        var config = JsonDocument.Parse(jsonConfig).RootElement;

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three check box form fields with distinct names.
        builder.Writeln("Form fields demonstration:");
        builder.Writeln();

        builder.Write("Option 1: ");
        FormField cb1 = builder.InsertCheckBox("CheckBox1", false, 15);
        builder.Writeln();

        builder.Write("Option 2: ");
        FormField cb2 = builder.InsertCheckBox("CheckBox2", false, 15);
        builder.Writeln();

        builder.Write("Option 3: ");
        FormField cb3 = builder.InsertCheckBox("CheckBox3", false, 15);
        builder.Writeln();

        // Access the collection of form fields in the document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Ensure that at least one form field exists.
        if (formFields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        // Update each check box according to the JSON configuration.
        foreach (FormField field in formFields)
        {
            // Only process check box fields.
            if (field.Type != FieldType.FieldFormCheckBox)
                continue;

            // Verify that the field has a name and that the configuration contains a value for it.
            if (string.IsNullOrEmpty(field.Name) || !config.TryGetProperty(field.Name, out JsonElement value))
                continue;

            // Set the Checked property based on the JSON boolean value.
            field.Checked = value.GetBoolean();

            // Optional validation: ensure the assignment succeeded.
            if (field.Checked != value.GetBoolean())
                throw new InvalidOperationException($"Failed to set the checked state for field '{field.Name}'.");
        }

        // Save the updated document.
        const string outputPath = "UpdatedFormFields.docx";
        doc.Save(outputPath);
    }
}

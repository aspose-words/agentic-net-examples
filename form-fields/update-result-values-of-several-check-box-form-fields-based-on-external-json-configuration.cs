using System;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to insert form fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three check box form fields with distinct names.
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

        // Parse the JSON configuration.
        using JsonDocument jsonDoc = JsonDocument.Parse(jsonConfig);
        JsonElement root = jsonDoc.RootElement;

        // Update each checkbox according to the JSON data.
        foreach (JsonProperty property in root.EnumerateObject())
        {
            string fieldName = property.Name;
            bool shouldBeChecked = property.Value.GetBoolean();

            // Retrieve the form field by name; throw if it does not exist.
            FormField formField = doc.Range.FormFields[fieldName];
            if (formField == null)
                throw new InvalidOperationException($"Form field '{fieldName}' not found.");

            // Ensure the field is a checkbox before setting the Checked property.
            if (formField.Type != FieldType.FieldFormCheckBox)
                throw new InvalidOperationException($"Form field '{fieldName}' is not a checkbox.");

            formField.Checked = shouldBeChecked;
        }

        // Update fields (not strictly required for checkboxes but follows best practice) and save the document.
        doc.UpdateFields();
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UpdatedFormFields.docx");
        doc.Save(outputPath);
    }
}

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
        // Create a new document and insert several checkbox form fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

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

        // Update each checkbox according to the configuration.
        foreach (KeyValuePair<string, bool> entry in config)
        {
            // Retrieve the form field by its name.
            FormField field = doc.Range.FormFields[entry.Key];
            if (field == null)
                throw new InvalidOperationException($"Form field '{entry.Key}' not found.");

            // Ensure the field is a checkbox before setting the Checked property.
            if (field.Type != FieldType.FieldFormCheckBox)
                throw new InvalidOperationException($"Form field '{entry.Key}' is not a checkbox.");

            field.Checked = entry.Value;
        }

        // Update fields (not strictly required for checkboxes but follows best practice) and save the document.
        doc.UpdateFields();
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UpdatedFormFields.docx");
        doc.Save(outputPath);
    }
}

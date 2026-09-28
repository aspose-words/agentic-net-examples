using System;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a document with three checkbox form fields.
        const string initialDocPath = "FormFields.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.InsertCheckBox("CheckBox1", false, 0);
        builder.Writeln(" CheckBox1 label");
        builder.InsertCheckBox("CheckBox2", true, 0);
        builder.Writeln(" CheckBox2 label");
        builder.InsertCheckBox("CheckBox3", false, 0);
        builder.Writeln(" CheckBox3 label");

        doc.Save(initialDocPath);

        // External JSON configuration that defines the desired checked state.
        string jsonConfig = @"{
            ""CheckBox1"": true,
            ""CheckBox2"": false,
            ""CheckBox3"": true
        }";

        // Parse JSON into a dictionary.
        Dictionary<string, bool> config = JsonSerializer.Deserialize<Dictionary<string, bool>>(jsonConfig) 
                                          ?? new Dictionary<string, bool>();

        // Load the document and update checkbox results according to the configuration.
        Document loadedDoc = new Document(initialDocPath);
        foreach (KeyValuePair<string, bool> kvp in config)
        {
            FormField field = loadedDoc.Range.FormFields[kvp.Key];
            if (field == null)
                throw new InvalidOperationException($"Form field '{kvp.Key}' not found.");

            if (field.Type != FieldType.FieldFormCheckBox)
                throw new InvalidOperationException($"Form field '{kvp.Key}' is not a checkbox.");

            field.Checked = kvp.Value;

            // Validate the update.
            if (field.Checked != kvp.Value)
                throw new InvalidOperationException($"Failed to set checkbox '{kvp.Key}' to '{kvp.Value}'.");
        }

        // Save the updated document.
        const string updatedDocPath = "FormFieldsUpdated.docx";
        loadedDoc.Save(updatedDocPath);

        Console.WriteLine($"Checkbox form fields have been updated and saved to '{updatedDocPath}'.");
    }
}

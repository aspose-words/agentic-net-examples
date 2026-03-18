using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Fields;

class UpdateCheckBoxFormFields
{
    static void Main()
    {
        // Paths to the input Word document, JSON configuration and the output document.
        string docPath = @"C:\Docs\InputDocument.docx";
        string jsonPath = @"C:\Docs\checkboxConfig.json";
        string outputPath = @"C:\Docs\UpdatedDocument.docx";

        // Load the JSON file which maps form field names to their desired checked state.
        // Example JSON content: { "CheckBox1": true, "CheckBox2": false }
        Dictionary<string, bool> checkBoxValues = JsonSerializer.Deserialize<Dictionary<string, bool>>(File.ReadAllText(jsonPath));

        // Load the Word document.
        Document doc = new Document(docPath);

        // Iterate through all form fields in the document.
        foreach (FormField field in doc.Range.FormFields)
        {
            // Process only check box form fields.
            if (field.Type == FieldType.FieldFormCheckBox)
            {
                // If the JSON configuration contains an entry for this field name, apply the value.
                if (checkBoxValues != null && checkBoxValues.TryGetValue(field.Name, out bool isChecked))
                {
                    field.Checked = isChecked;
                }
            }
        }

        // Save the modified document.
        doc.Save(outputPath);
    }
}

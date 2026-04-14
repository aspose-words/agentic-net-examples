using System;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Fields;

namespace FormFieldUpdater
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document and a builder to insert form fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert several check box form fields with distinct names.
            builder.Write("Option A: ");
            FormField checkBoxA = builder.InsertCheckBox("OptionA", false, 0);
            builder.InsertBreak(BreakType.ParagraphBreak);

            builder.Write("Option B: ");
            FormField checkBoxB = builder.InsertCheckBox("OptionB", false, 0);
            builder.InsertBreak(BreakType.ParagraphBreak);

            builder.Write("Option C: ");
            FormField checkBoxC = builder.InsertCheckBox("OptionC", false, 0);
            builder.InsertBreak(BreakType.ParagraphBreak);

            // Ensure that at least one form field exists.
            if (doc.Range.FormFields.Count == 0)
                throw new InvalidOperationException("The document does not contain any form fields.");

            // External JSON configuration that maps field names to their desired checked state.
            string jsonConfig = @"{
                ""OptionA"": true,
                ""OptionB"": false,
                ""OptionC"": true
            }";

            // Parse the JSON into a dictionary.
            Dictionary<string, bool> config = JsonSerializer.Deserialize<Dictionary<string, bool>>(jsonConfig)
                ?? throw new InvalidOperationException("Failed to parse JSON configuration.");

            // Update each check box according to the configuration.
            foreach (KeyValuePair<string, bool> entry in config)
            {
                // Retrieve the form field by name; the indexer returns null if not found.
                FormField field = doc.Range.FormFields[entry.Key];
                if (field == null)
                    throw new KeyNotFoundException($"Form field '{entry.Key}' was not found in the document.");

                // Verify that the field is a check box.
                if (field.Type != FieldType.FieldFormCheckBox)
                    throw new InvalidOperationException($"Form field '{entry.Key}' is not a check box.");

                // Set the checked status.
                field.Checked = entry.Value;

                // Validate that the assignment succeeded.
                if (field.Checked != entry.Value)
                    throw new InvalidOperationException($"Failed to set the checked state for field '{entry.Key}'.");
            }

            // Update fields (not strictly necessary for check boxes but follows best practice).
            doc.UpdateFields();

            // Save the updated document.
            const string outputPath = "UpdatedFormFields.docx";
            doc.Save(outputPath);
        }
    }
}

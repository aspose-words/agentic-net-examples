using System;
using Aspose.Words;
using Aspose.Words.Fields;

class FormatFormFields
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\InputDocument.docx";

        // Path where the formatted document will be saved.
        string outputPath = @"C:\Docs\FormattedDocument.docx";

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Optional: turn off the default gray shading for form fields.
        doc.ShadeFormData = false;

        // Iterate through all form fields in the document.
        foreach (FormField field in doc.Range.FormFields)
        {
            // Enable the field (make it editable).
            field.Enabled = true;

            // Example: set a help text for each field.
            field.HelpText = "Please fill out this field.";

            // Apply field‑type‑specific formatting.
            switch (field.Type)
            {
                case FieldType.FieldFormCheckBox:
                    // Ensure the checkbox is unchecked by default.
                    field.Checked = false;
                    field.Default = false;
                    break;

                case FieldType.FieldFormDropDown:
                    // Select the first item in the drop‑down list.
                    if (field.DropDownItems.Count > 0)
                        field.DropDownSelectedIndex = 0;
                    break;

                case FieldType.FieldFormTextInput:
                    // Set placeholder text for text input fields.
                    field.TextInputDefault = "Enter text here";
                    // Optionally set a maximum length.
                    field.MaxLength = 100;
                    break;
            }
        }

        // Update all fields so that any changes are reflected in the document.
        doc.UpdateFields();

        // Save the modified document.
        doc.Save(outputPath);
    }
}

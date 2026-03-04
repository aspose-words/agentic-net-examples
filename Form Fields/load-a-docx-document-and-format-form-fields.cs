using System;
using Aspose.Words;
using Aspose.Words.Fields;

class FormatFormFields
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "input.docx";

        // Path where the modified document will be saved.
        string outputPath = "output.docx";

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Turn off the default gray shading for form fields (optional).
        doc.ShadeFormData = false;

        // Iterate through all form fields in the document.
        foreach (FormField field in doc.Range.FormFields)
        {
            // Enable the field so the user can interact with it.
            field.Enabled = true;

            // Set a generic help text for demonstration purposes.
            field.HelpText = "Enter value";

            // Apply specific formatting based on the field type.
            switch (field.Type)
            {
                case FieldType.FieldFormCheckBox:
                    // Ensure the checkbox is unchecked by default.
                    field.Checked = false;
                    field.Default = false;
                    break;

                case FieldType.FieldFormDropDown:
                    // Select the first item in the drop‑down list.
                    field.DropDownSelectedIndex = 0;
                    break;

                case FieldType.FieldFormTextInput:
                    // Set a placeholder text for text input fields.
                    field.TextInputDefault = "Placeholder";
                    // Optionally set a maximum length.
                    field.MaxLength = 100;
                    break;
            }
        }

        // Update all fields to reflect any changes made programmatically.
        doc.UpdateFields();

        // Save the modified document.
        doc.Save(outputPath);
    }
}

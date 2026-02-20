using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Loading; // Added for LoadOptions

class Program
{
    static void Main()
    {
        // Load the DOCX document with explicit load options.
        var loadOptions = new LoadOptions { LoadFormat = LoadFormat.Docx };
        Document doc = new Document("input.docx", loadOptions);

        // Apply gray shading to all form fields.
        doc.ShadeFormData = true;

        // Iterate through each form field in the document.
        foreach (FormField field in doc.Range.FormFields)
        {
            // Make sure the field is enabled for user interaction.
            field.Enabled = true;

            // Provide a default help text if none is set.
            if (string.IsNullOrEmpty(field.HelpText))
                field.HelpText = "Please fill out this field";

            // For text input fields, set a placeholder value when empty.
            if (field.Type == FieldType.FieldFormTextInput && string.IsNullOrEmpty(field.Result))
                field.Result = "Enter value";

            // For check box fields, ensure the default state is unchecked.
            if (field.Type == FieldType.FieldFormCheckBox)
                field.Checked = false; // Correct property for check box state
        }

        // Save the modified document.
        doc.Save("output.docx");
    }
}

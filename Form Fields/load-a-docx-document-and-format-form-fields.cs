using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the existing DOCX document from the file system.
        Document doc = new Document("Input.docx");

        // Disable the default gray shading that appears on form fields.
        doc.ShadeFormData = false;

        // Iterate through every form field in the document.
        foreach (FormField field in doc.Range.FormFields)
        {
            // Ensure the field is enabled so the user can interact with it.
            field.Enabled = true;

            // Provide a generic help tooltip for the field.
            field.HelpText = "Please fill out this field.";

            // Apply type‑specific formatting or default values.
            switch (field.Type)
            {
                case FieldType.FieldFormTextInput:
                    // For text input fields, set a sample placeholder text.
                    field.Result = "Sample text";
                    break;

                case FieldType.FieldFormCheckBox:
                    // For check boxes, make sure they are unchecked by default.
                    field.Checked = false;
                    break;

                case FieldType.FieldFormDropDown:
                    // For drop‑down fields, select the first item in the list.
                    field.DropDownSelectedIndex = 0;
                    break;
            }
        }

        // Recalculate all fields to reflect the changes made programmatically.
        doc.UpdateFields();

        // Save the modified document to a new file.
        doc.Save("Output.docx");
    }
}

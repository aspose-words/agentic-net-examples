using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and insert a checkbox form field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Write("Toggle this check box: ");
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 20);
        doc.Save("FormWithCheckBox.docx");

        // Read external configuration (environment variable).
        string env = Environment.GetEnvironmentVariable("CHECKBOX_CHECKED");
        bool configChecked = false;
        if (!string.IsNullOrEmpty(env) && bool.TryParse(env, out bool parsed))
        {
            configChecked = parsed;
        }

        // Locate the checkbox by name.
        FormField targetField = doc.Range.FormFields["MyCheckBox"];
        if (targetField == null)
        {
            throw new InvalidOperationException("Checkbox form field 'MyCheckBox' not found.");
        }

        // Ensure the field is a checkbox.
        if (targetField.Type != FieldType.FieldFormCheckBox)
        {
            throw new InvalidOperationException("Form field 'MyCheckBox' is not a checkbox.");
        }

        // Set the checked state based on the configuration.
        targetField.Checked = configChecked;

        // Validate the update.
        if (targetField.Checked != configChecked)
        {
            throw new InvalidOperationException("Failed to set the checkbox state.");
        }

        // Save the modified document.
        doc.Save("FormWithCheckBox_Toggled.docx");
    }
}

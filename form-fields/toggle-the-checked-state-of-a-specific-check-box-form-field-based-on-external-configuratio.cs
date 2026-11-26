using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a checkbox form field with a known name.
        // Parameters: name, isChecked (default), size (in points).
        builder.InsertCheckBox("MyCheckBox", false, 12);

        // Save the initial document (optional, shows the created file).
        doc.Save("sample.docx");

        // Retrieve the checkbox form field from the document.
        var checkBoxField = doc.Range.FormFields["MyCheckBox"];
        if (checkBoxField == null)
        {
            throw new InvalidOperationException("Checkbox form field 'MyCheckBox' was not found.");
        }

        // Read external configuration from an environment variable.
        // Expected values: "true" or "false". Any other value will cause a toggle.
        string envValue = Environment.GetEnvironmentVariable("CHECKBOX_STATE");
        bool setChecked;
        if (bool.TryParse(envValue, out bool parsed))
        {
            setChecked = parsed; // Use the configured value.
        }
        else
        {
            // If the configuration is missing or invalid, toggle the current state.
            setChecked = !checkBoxField.Checked;
        }

        // Update the checkbox state.
        checkBoxField.Checked = setChecked;

        // Validate that the update succeeded.
        if (checkBoxField.Checked != setChecked)
        {
            throw new InvalidOperationException("Failed to update the checkbox state.");
        }

        // Save the modified document.
        doc.Save("output.docx");
    }
}

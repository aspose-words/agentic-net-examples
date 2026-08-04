using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a checkbox form field with a known name.
        const string checkBoxName = "MyCheckBox";
        builder.Write("Toggle this check box: ");
        FormField insertedCheckBox = builder.InsertCheckBox(checkBoxName, false, 0);
        if (insertedCheckBox == null)
            throw new InvalidOperationException("Failed to create the checkbox form field.");

        // Read external configuration (environment variable "CHECKBOX_STATE").
        // Expected values: "true" or "false". Default is false if not set or invalid.
        string envValue = Environment.GetEnvironmentVariable("CHECKBOX_STATE");
        bool shouldBeChecked = false;
        if (!string.IsNullOrEmpty(envValue) && bool.TryParse(envValue, out bool parsed))
            shouldBeChecked = parsed;

        // Locate the checkbox by name in the document's form fields collection.
        FormField targetField = null;
        foreach (FormField field in doc.Range.FormFields)
        {
            if (field.Name == checkBoxName)
            {
                targetField = field;
                break;
            }
        }

        if (targetField == null)
            throw new InvalidOperationException($"Form field '{checkBoxName}' not found.");

        // Toggle the checked state based on the external configuration.
        targetField.Checked = shouldBeChecked;

        // Validate that the state was applied.
        if (targetField.Checked != shouldBeChecked)
            throw new InvalidOperationException("Failed to set the checkbox state.");

        // Save the modified document.
        const string outputPath = "ToggleCheckbox.docx";
        doc.Save(outputPath);
    }
}

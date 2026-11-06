using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to add a checkbox form field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a checkbox named "AcceptTerms". Default state is unchecked.
        builder.InsertCheckBox("AcceptTerms", false, 0);

        // Save the initial document.
        const string initialPath = "FormFieldsExample.docx";
        doc.Save(initialPath);

        // ----- External configuration -----
        // Read the desired checkbox state from an environment variable.
        // Expected values: "true" or "false". If missing or invalid, default to false.
        string envValue = Environment.GetEnvironmentVariable("CHECKBOX_STATE");
        bool configChecked = false;
        if (!string.IsNullOrEmpty(envValue) && bool.TryParse(envValue, out bool parsed))
        {
            configChecked = parsed;
        }

        // Load the document that contains the form field.
        Document loadedDoc = new Document(initialPath);

        // Access the collection of form fields.
        FormFieldCollection fields = loadedDoc.Range.FormFields;

        // Validate that the expected checkbox exists.
        FormField checkBoxField = fields["AcceptTerms"];
        if (checkBoxField == null)
        {
            throw new InvalidOperationException("Checkbox form field 'AcceptTerms' not found.");
        }

        // Ensure the field is indeed a checkbox.
        if (checkBoxField.Type != FieldType.FieldFormCheckBox)
        {
            throw new InvalidOperationException("Form field 'AcceptTerms' is not a checkbox.");
        }

        // Update the checked state based on the external configuration.
        checkBoxField.Checked = configChecked;

        // Verify that the update succeeded.
        if (checkBoxField.Checked != configChecked)
        {
            throw new InvalidOperationException("Failed to update the checkbox state.");
        }

        // Save the modified document.
        const string outputPath = "FormFieldsExample_Updated.docx";
        loadedDoc.Save(outputPath);
    }
}

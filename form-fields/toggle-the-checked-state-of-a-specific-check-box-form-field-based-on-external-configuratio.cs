using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and insert a checkbox form field named "MyCheckBox".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertCheckBox("MyCheckBox", false, 0);
        // Save the initial document.
        doc.Save("initial.docx");

        // Load the document to simulate a separate operation.
        Document loadedDoc = new Document("initial.docx");

        // Read external configuration: environment variable "TOGGLE_CHECKBOX".
        // If the variable is set to "true" (case‑insensitive), the checkbox will be toggled.
        string toggleSetting = Environment.GetEnvironmentVariable("TOGGLE_CHECKBOX");
        bool shouldToggle = string.Equals(toggleSetting, "true", StringComparison.OrdinalIgnoreCase);

        // Retrieve the checkbox form field by name.
        FormField checkBoxField = loadedDoc.Range.FormFields["MyCheckBox"];
        if (checkBoxField == null)
        {
            throw new InvalidOperationException("Checkbox form field 'MyCheckBox' was not found.");
        }

        // Verify that the field is indeed a checkbox.
        if (checkBoxField.Type != FieldType.FieldFormCheckBox)
        {
            throw new InvalidOperationException("Form field 'MyCheckBox' is not a checkbox.");
        }

        // Toggle the checked state if required.
        if (shouldToggle)
        {
            checkBoxField.Checked = !checkBoxField.Checked;
        }

        // Save the modified document.
        loadedDoc.Save("output.docx");
    }
}

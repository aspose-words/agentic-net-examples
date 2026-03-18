using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the document that already contains the checkbox form field.
        Document doc = new Document("Input.docx");

        // Retrieve the desired checked state from an external source.
        // Replace this stub with actual configuration logic as needed.
        bool targetCheckedState = GetDesiredCheckedState();

        // Locate the checkbox by its name (bookmark name) in the form fields collection.
        FormField checkBox = doc.Range.FormFields["MyCheckBox"];
        if (checkBox != null && checkBox.Type == FieldType.FieldFormCheckBox)
        {
            // Apply the configuration value to the checkbox.
            checkBox.Checked = targetCheckedState;
        }

        // Save the updated document.
        doc.Save("Output.docx");
    }

    // Example placeholder for external configuration retrieval.
    static bool GetDesiredCheckedState()
    {
        // Example: read an environment variable named "CHECKBOX_STATE".
        // Returns true if the variable is set to "true" (case‑insensitive), otherwise false.
        string envValue = Environment.GetEnvironmentVariable("CHECKBOX_STATE");
        return string.Equals(envValue, "true", StringComparison.OrdinalIgnoreCase);
    }
}

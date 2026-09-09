using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    // Entry point of the console application.
    public static void Main()
    {
        // Path for the initial document and the updated document.
        const string initialDocPath = "FormFields.docx";
        const string updatedDocPath = "FormFields_Updated.docx";

        // -----------------------------------------------------------------
        // 1. Create a new document and insert a checkbox form field.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with a checkbox named "MyCheckBox".
        builder.Writeln("Toggle this checkbox based on configuration:");
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 0);
        // Optional: set a readable size for the checkbox.
        checkBox.IsCheckBoxExactSize = true;
        checkBox.CheckBoxSize = 12.0;

        // Save the document that contains the form field.
        doc.Save(initialDocPath);

        // -----------------------------------------------------------------
        // 2. Load the document (simulating a separate operation) and
        //    update the checkbox state according to external configuration.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(initialDocPath);
        FormFieldCollection formFields = loadedDoc.Range.FormFields;

        // Validate that the expected form field exists.
        FormField? targetField = formFields["MyCheckBox"];
        if (targetField == null)
            throw new InvalidOperationException("Form field 'MyCheckBox' was not found.");

        // Read external configuration. Here we use an environment variable.
        // Expected values: "true" or "false" (case‑insensitive). Default is false.
        bool desiredState = GetDesiredCheckedStateFromEnv();

        // Toggle the checkbox state.
        targetField.Checked = desiredState;

        // Save the updated document.
        loadedDoc.Save(updatedDocPath);
    }

    // Reads the environment variable "CHECKBOX_CHECKED" and converts it to a bool.
    // Returns false if the variable is missing or cannot be parsed.
    private static bool GetDesiredCheckedStateFromEnv()
    {
        string? envValue = Environment.GetEnvironmentVariable("CHECKBOX_CHECKED");
        if (string.IsNullOrWhiteSpace(envValue))
            return false;

        return bool.TryParse(envValue.Trim(), out bool result) && result;
    }
}

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
        FormField checkBox = builder.InsertCheckBox(checkBoxName, false, 0);
        builder.InsertParagraph();

        // Read external configuration (environment variable) that determines the desired state.
        // Expected values: "true" or "false". Default to false if not set or invalid.
        string configValue = Environment.GetEnvironmentVariable("CHECKBOX_STATE");
        bool desiredState = false;
        if (!string.IsNullOrWhiteSpace(configValue) && bool.TryParse(configValue, out bool parsed))
        {
            desiredState = parsed;
        }

        // Locate the checkbox form field by name and validate its existence.
        FormField? targetField = doc.Range.FormFields[checkBoxName];
        if (targetField == null)
        {
            throw new InvalidOperationException($"Form field '{checkBoxName}' was not found in the document.");
        }

        // Toggle the checked state according to the external configuration.
        targetField.Checked = desiredState;

        // Save the modified document.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}

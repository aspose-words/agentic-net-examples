using System;
using Aspose.Words;
using Aspose.Words.Fields; // Required for FormField

public class Program
{
    public static void Main()
    {
        // Simulated external configuration: desired checked state for the checkbox.
        bool configChecked = true; // This could be read from a file, environment variable, etc.

        // -----------------------------------------------------------------
        // Create a new document and insert a checkbox form field.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Write("Please tick the box: ");

        // Insert a checkbox named "MyCheckBox" with an initial unchecked state.
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 0);

        // Save the initial document.
        doc.Save("FormFields.docx");

        // -----------------------------------------------------------------
        // Load the document and toggle the checkbox based on the configuration.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document("FormFields.docx");

        // Retrieve the checkbox form field by its name.
        FormField field = loadedDoc.Range.FormFields["MyCheckBox"];
        if (field == null)
            throw new InvalidOperationException("Form field 'MyCheckBox' was not found in the document.");

        // Apply the configuration value to the Checked property.
        field.Checked = configChecked;

        // Validate that the property was set correctly.
        if (field.Checked != configChecked)
            throw new InvalidOperationException("Failed to update the checked state of the form field.");

        // Save the updated document.
        loadedDoc.Save("FormFields_Updated.docx");
    }
}

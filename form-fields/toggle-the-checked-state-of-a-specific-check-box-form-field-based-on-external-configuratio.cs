using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Simulated external configuration that determines the desired checkbox state.
        bool configShouldCheck = true; // This could be read from a config file, environment variable, etc.

        // Create a new blank document and a builder to insert content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a description and a legacy checkbox form field with a known name.
        builder.Write("Please indicate your agreement: ");
        // InsertCheckBox(name, defaultChecked, size). Size 0 lets Word choose the size automatically.
        builder.InsertCheckBox("AgreementCheckBox", false, 0);

        // Validate that the document contains at least one form field.
        FormFieldCollection formFields = doc.Range.FormFields;
        if (formFields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        // Retrieve the checkbox by its name.
        FormField checkBoxField = doc.Range.FormFields["AgreementCheckBox"];
        if (checkBoxField == null)
            throw new InvalidOperationException("The expected checkbox form field 'AgreementCheckBox' was not found.");

        // Toggle the checked state based on the external configuration.
        checkBoxField.Checked = configShouldCheck;

        // Verify that the assignment succeeded.
        if (checkBoxField.Checked != configShouldCheck)
            throw new InvalidOperationException("Failed to set the checkbox state as per configuration.");

        // Save the modified document.
        doc.Save("ToggledCheckbox.docx");
    }
}

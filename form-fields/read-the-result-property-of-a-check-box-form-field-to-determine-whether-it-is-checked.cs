using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a checkbox form field with a name and default unchecked state.
        // Parameters: field name, isChecked, size (points).
        builder.InsertCheckBox("SampleCheckBox", false, 50);

        // Save the document that contains the form field.
        const string fileName = "SampleForm.docx";
        doc.Save(fileName);

        // Load the document (could also reuse the same instance).
        Document loadedDoc = new Document(fileName);

        // Retrieve the checkbox form field by name.
        FormField checkBoxField = loadedDoc.Range.FormFields["SampleCheckBox"];
        if (checkBoxField == null)
        {
            throw new InvalidOperationException("The expected checkbox form field was not found.");
        }

        // Read the checked state using the Checked property (preferred for checkboxes).
        bool isChecked = checkBoxField.Checked;

        // Output the result.
        Console.WriteLine($"Checkbox '{checkBoxField.Name}' is checked: {isChecked}");

        // Optionally, modify the checkbox state and save again.
        checkBoxField.Checked = true;
        loadedDoc.Save("SampleForm_Modified.docx");
    }
}

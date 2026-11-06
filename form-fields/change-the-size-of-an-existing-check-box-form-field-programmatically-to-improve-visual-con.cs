using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a legacy checkbox form field named "CheckBox1".
        // Parameters: name, isChecked (false), size (default 0 points – will be changed later).
        builder.InsertCheckBox("CheckBox1", false, 0);

        // Save the initial document containing the checkbox.
        const string initialPath = "CheckboxFormField.docx";
        doc.Save(initialPath);

        // Load the document back (demonstrates load step).
        Document loadedDoc = new Document(initialPath);

        // Retrieve the checkbox form field by name.
        var checkBoxField = loadedDoc.Range.FormFields["CheckBox1"];
        if (checkBoxField == null)
            throw new InvalidOperationException("The checkbox form field 'CheckBox1' was not found.");

        // Change the visual size of the checkbox to 20 points for consistency.
        checkBoxField.CheckBoxSize = 20;

        // Save the modified document.
        const string outputPath = "CheckboxFormField_Modified.docx";
        loadedDoc.Save(outputPath);
    }
}

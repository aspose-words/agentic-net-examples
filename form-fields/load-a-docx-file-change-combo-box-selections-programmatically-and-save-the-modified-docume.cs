using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        const string originalPath = "ComboBoxForm.docx";
        const string modifiedPath = "ComboBoxForm_Modified.docx";
        const string comboName = "MyCombo";

        // -----------------------------------------------------------------
        // Create a new document and insert a combo box (drop‑down) form field.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a combo box with three items and select the first one by default.
        builder.InsertComboBox(comboName, new[] { "Option1", "Option2", "Option3" }, 0);

        // Save the document that now contains the form field.
        doc.Save(originalPath);

        // ---------------------------------------------------------------
        // Load the document from disk and modify the combo box selection.
        // ---------------------------------------------------------------
        Document loadedDoc = new Document(originalPath);

        // Retrieve the combo box form field by its name.
        FormField comboField = loadedDoc.Range.FormFields[comboName];
        if (comboField == null)
        {
            throw new InvalidOperationException($"Form field '{comboName}' was not found in the document.");
        }

        // Change the selected value of the combo box.
        comboField.Result = "Option3";

        // Verify that the value was updated correctly.
        if (comboField.Result != "Option3")
        {
            throw new InvalidOperationException("Failed to update the combo box selection.");
        }

        // Save the modified document.
        loadedDoc.Save(modifiedPath);
    }
}

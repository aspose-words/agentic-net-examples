using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Paths for the original and modified documents.
        const string originalPath = "FormFields.docx";
        const string modifiedPath = "FormFields_Modified.docx";

        // -------------------------------------------------
        // 1. Create a sample DOCX with a combo box field.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some prompt text.
        builder.Write("Pick a fruit: ");

        // Define the items for the combo box and insert it.
        string[] items = { "Apple", "Banana", "Cherry" };
        // InsertComboBox creates a legacy combo box form field.
        builder.InsertComboBox("FruitCombo", items, 0); // Default selection is the first item (Apple).

        // Save the document that contains the form field.
        doc.Save(originalPath);

        // -------------------------------------------------
        // 2. Load the document and modify the combo box selection.
        // -------------------------------------------------
        Document loadedDoc = new Document(originalPath);

        // Access the collection of form fields.
        FormFieldCollection formFields = loadedDoc.Range.FormFields;

        // Retrieve the combo box by its name.
        FormField comboBox = formFields["FruitCombo"];
        if (comboBox == null)
            throw new InvalidOperationException("The combo box 'FruitCombo' was not found in the document.");

        // Validate that the field is indeed a drop‑down.
        if (comboBox.Type != FieldType.FieldFormDropDown)
            throw new InvalidOperationException("The field 'FruitCombo' is not a drop‑down form field.");

        // Change the selected item to "Banana" (index 1).
        comboBox.DropDownSelectedIndex = 1;

        // Alternatively, you could set the Result property:
        // comboBox.Result = "Banana";

        // -------------------------------------------------
        // 3. Save the modified document.
        // -------------------------------------------------
        loadedDoc.Save(modifiedPath);
    }
}

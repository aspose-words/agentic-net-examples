using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary input and final output documents.
        const string inputPath = "input.docx";
        const string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // 1. Create a DOCX file with a combo box form field.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a prompt and insert a combo box named "FruitCombo".
        builder.Write("Pick a fruit: ");
        string[] items = { "Apple", "Banana", "Cherry" };
        // InsertComboBox creates the form field and a bookmark with the same name.
        builder.InsertComboBox("FruitCombo", items, 0); // default selection is the first item.

        // Save the document that will be loaded later.
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the existing DOCX file.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // -----------------------------------------------------------------
        // 3. Locate the combo box form field and change its selection.
        // -----------------------------------------------------------------
        FormField comboBox = loadedDoc.Range.FormFields["FruitCombo"];
        if (comboBox == null)
            throw new InvalidOperationException("The combo box 'FruitCombo' was not found in the document.");

        // Validate that the field is indeed a drop‑down.
        if (comboBox.Type != FieldType.FieldFormDropDown)
            throw new InvalidOperationException("The field 'FruitCombo' is not a drop‑down form field.");

        // Change the selected index to 1 (second item: "Banana").
        comboBox.DropDownSelectedIndex = 1;

        // Optional validation: ensure the Result reflects the new selection.
        if (!comboBox.Result.Equals("Banana", StringComparison.Ordinal))
            throw new InvalidOperationException("Failed to update the combo box selection.");

        // -----------------------------------------------------------------
        // 4. Save the modified document.
        // -----------------------------------------------------------------
        loadedDoc.Save(outputPath);
    }
}

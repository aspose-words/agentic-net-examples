using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample DOCX file that contains a combo box form field.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a prompt and insert a combo box with three items.
        builder.Writeln("Select a fruit:");
        string[] items = { "Apple", "Banana", "Cherry" };
        // InsertComboBox(name, items, selectedIndex) – default selection is the first item (Apple).
        FormField comboBox = builder.InsertComboBox("FruitCombo", items, 0);

        // Save the document that will be loaded later.
        string inputPath = "FormFieldsInput.docx";
        doc.Save(inputPath);

        // ---------------------------------------------------------------
        // 2. Load the DOCX file, modify the combo box selection programmatically.
        // ---------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // Access the collection of form fields.
        FormFieldCollection formFields = loadedDoc.Range.FormFields;

        // Validate that the expected combo box exists.
        FormField fruitCombo = formFields["FruitCombo"];
        if (fruitCombo == null)
            throw new InvalidOperationException("Combo box 'FruitCombo' was not found in the document.");

        // Change the selected index to 2 (third item – "Cherry").
        fruitCombo.DropDownSelectedIndex = 2;

        // Verify that the selection was updated correctly.
        if (fruitCombo.DropDownSelectedIndex != 2)
            throw new InvalidOperationException("Failed to update the combo box selection.");

        // ---------------------------------------------------------------
        // 3. Save the modified document.
        // ---------------------------------------------------------------
        string outputPath = "FormFieldsOutput.docx";
        loadedDoc.Save(outputPath);
    }
}

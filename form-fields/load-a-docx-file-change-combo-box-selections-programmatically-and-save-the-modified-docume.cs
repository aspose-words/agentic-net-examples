using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Paths for the original and modified documents.
        const string inputPath = "FormFields.docx";
        const string outputPath = "FormFields_Modified.docx";

        // -------------------------------------------------
        // Create a sample DOCX file that contains a combo box.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a prompt and insert a combo box with three items.
        builder.Write("Select a fruit: ");
        string[] items = { "Apple", "Banana", "Cherry" };
        // InsertComboBox(name, items, selectedIndex) – creates the form field.
        builder.InsertComboBox("FruitCombo", items, 0); // default selection is the first item.

        // Save the document that will be loaded later.
        doc.Save(inputPath);

        // -------------------------------------------------
        // Load the document, modify the combo box selection, and save.
        // -------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // Access the combo box by its name via the FormFields collection.
        FormField comboBox = loadedDoc.Range.FormFields["FruitCombo"];
        if (comboBox == null)
            throw new InvalidOperationException("The combo box 'FruitCombo' was not found in the document.");

        // Change the selected item. Index is zero‑based, so 2 selects "Cherry".
        comboBox.DropDownSelectedIndex = 2;

        // Optionally, verify the change.
        if (comboBox.DropDownSelectedIndex != 2)
            throw new InvalidOperationException("Failed to update the combo box selection.");

        // Save the modified document.
        loadedDoc.Save(outputPath);
    }
}

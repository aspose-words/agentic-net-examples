using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Define file names.
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "FormFields.docx");
        string modifiedPath = Path.Combine(Directory.GetCurrentDirectory(), "FormFields_Modified.docx");

        // -----------------------------------------------------------------
        // 1. Create a DOCX file with a combo box form field.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a prompt.
        builder.Write("Select a fruit: ");

        // Insert a combo box named "FruitCombo" with three items.
        string[] items = { "Apple", "Banana", "Cherry" };
        FormField comboBox = builder.InsertComboBox("FruitCombo", items, 0); // 0 = Apple selected by default.

        // Save the document that now contains the form field.
        doc.Save(originalPath);

        // -----------------------------------------------------------------
        // 2. Load the previously saved DOCX file.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(originalPath);

        // -----------------------------------------------------------------
        // 3. Locate the combo box and change its selected item.
        // -----------------------------------------------------------------
        FormFieldCollection formFields = loadedDoc.Range.FormFields;

        // Validate that the expected form field exists.
        FormField fruitField = formFields["FruitCombo"];
        if (fruitField == null)
            throw new InvalidOperationException("The combo box 'FruitCombo' was not found in the document.");

        // Change the selected index to 1 (Banana). Index is zero‑based.
        fruitField.DropDownSelectedIndex = 1;

        // Optional validation to ensure the change was applied.
        if (fruitField.DropDownSelectedIndex != 1)
            throw new InvalidOperationException("Failed to update the combo box selection.");

        // -----------------------------------------------------------------
        // 4. Save the modified document.
        // -----------------------------------------------------------------
        loadedDoc.Save(modifiedPath);
    }
}

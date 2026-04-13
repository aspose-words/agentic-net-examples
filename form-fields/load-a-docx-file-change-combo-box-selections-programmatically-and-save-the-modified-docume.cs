using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string originalPath = Path.Combine(Environment.CurrentDirectory, "FormFields.docx");
        string modifiedPath = Path.Combine(Environment.CurrentDirectory, "FormFields_Modified.docx");

        // -----------------------------------------------------------------
        // 1. Create a new document and insert a combo box form field.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Write("Select a fruit: ");

        // Items for the combo box.
        string[] items = { "Apple", "Banana", "Cherry" };

        // Insert the combo box with the first item selected (index 0).
        FormField comboBox = builder.InsertComboBox("FruitCombo", items, 0);

        // Save the document that now contains the form field.
        doc.Save(originalPath);

        // -----------------------------------------------------------------
        // 2. Load the document, modify the combo box selection, and save.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(originalPath);

        // Access the collection of form fields.
        FormFieldCollection formFields = loadedDoc.Range.FormFields;

        // Retrieve the combo box by its name.
        FormField fruitCombo = formFields["FruitCombo"];
        if (fruitCombo == null)
            throw new InvalidOperationException("The expected combo box 'FruitCombo' was not found.");

        // Change the selected index to the third item ("Cherry").
        fruitCombo.DropDownSelectedIndex = 2;

        // Validate that the result reflects the new selection.
        if (!string.Equals(fruitCombo.Result, "Cherry", StringComparison.Ordinal))
            throw new InvalidOperationException("Failed to update the combo box selection.");

        // Save the modified document.
        loadedDoc.Save(modifiedPath);
    }
}

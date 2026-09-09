using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    // Reusable method that inserts a combo box form field.
    // Parameters:
    //   builder      - DocumentBuilder positioned where the field should be inserted.
    //   name         - Name of the form field (bookmark will be created automatically).
    //   items        - Array of strings that will appear in the drop‑down list.
    //   defaultIndex - Zero‑based index of the item that should be selected by default.
    // Returns the inserted FormField instance.
    public static FormField AddComboBox(DocumentBuilder builder, string name, string[] items, int defaultIndex)
    {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        if (items == null) throw new ArgumentNullException(nameof(items));
        if (defaultIndex < 0 || defaultIndex >= items.Length)
            throw new ArgumentOutOfRangeException(nameof(defaultIndex), "Default index must be within the items array.");

        // Insert the combo box using the Aspose.Words API.
        FormField comboBox = builder.InsertComboBox(name, items, defaultIndex);
        return comboBox;
    }

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a prompt before the combo box.
        builder.Write("Please select a fruit: ");

        // Define the items for the combo box.
        string[] fruitItems = { "Apple", "Banana", "Cherry", "Date" };

        // Insert the combo box with "Banana" selected by default (index 1).
        FormField fruitCombo = AddComboBox(builder, "FruitCombo", fruitItems, 1);

        // Optionally, demonstrate accessing the field after insertion.
        // Verify that the selected item matches the default index.
        if (fruitCombo.DropDownSelectedIndex != 1 || fruitCombo.Result != "Banana")
            throw new InvalidOperationException("Combo box was not initialized correctly.");

        // Save the document to disk.
        doc.Save("ComboBoxFormField.docx");
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    // Reusable method that inserts a combo box form field.
    // Parameters:
    //   builder      - DocumentBuilder positioned where the field should be inserted.
    //   name         - Name of the form field (also creates a bookmark with the same name).
    //   items        - Array of strings that will appear in the drop‑down list.
    //   defaultIndex - Zero‑based index of the item that will be selected by default.
    // Returns the inserted FormField instance.
    public static FormField AddComboBoxFormField(DocumentBuilder builder, string name, string[] items, int defaultIndex)
    {
        // Validate arguments to avoid runtime errors.
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        if (string.IsNullOrEmpty(name)) throw new ArgumentException("Form field name cannot be null or empty.", nameof(name));
        if (items == null || items.Length == 0) throw new ArgumentException("Items collection cannot be null or empty.", nameof(items));
        if (defaultIndex < 0 || defaultIndex >= items.Length) throw new ArgumentOutOfRangeException(nameof(defaultIndex));

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
        builder.Write("Pick a fruit: ");

        // Define the items for the combo box.
        string[] fruitItems = { "Apple", "Banana", "Cherry", "Date" };

        // Insert the combo box with "Banana" selected by default (index 1).
        FormField fruitCombo = AddComboBoxFormField(builder, "FruitCombo", fruitItems, 1);

        // Optional: verify that the default selected item matches the expected value.
        if (fruitCombo.Result != fruitItems[1])
            throw new InvalidOperationException("The combo box default selection was not set correctly.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComboBoxFormField.docx");
        doc.Save(outputPath);
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    /// <summary>
    /// Adds a combo box (drop‑down) form field to the document at the builder's current position.
    /// </summary>
    /// <param name="builder">The DocumentBuilder positioned where the field should be inserted.</param>
    /// <param name="name">The name of the form field (also creates a bookmark with the same name).</param>
    /// <param name="items">Array of strings that will appear as selectable items.</param>
    /// <param name="defaultIndex">
    /// Zero‑based index of the item that should be selected by default.
    /// If the index is out of range it will be clamped to a valid value.
    /// </param>
    /// <returns>The inserted FormField instance.</returns>
    public static FormField AddComboBoxFormField(DocumentBuilder builder, string name, string[] items, int defaultIndex)
    {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        if (items == null) throw new ArgumentNullException(nameof(items));
        if (items.Length == 0) throw new ArgumentException("The items collection must contain at least one entry.", nameof(items));

        // Ensure the default index is within the bounds of the items array.
        int safeIndex = Math.Max(0, Math.Min(defaultIndex, items.Length - 1));

        // Insert the combo box using the Aspose.Words API as required by the rules.
        FormField comboBox = builder.InsertComboBox(name, items, safeIndex);
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

        // Add the combo box with "Banana" selected by default (index 1).
        FormField fruitCombo = AddComboBoxFormField(builder, "FruitCombo", fruitItems, 1);

        // Optional: verify that the field was added correctly.
        if (fruitCombo == null || fruitCombo.Result != fruitItems[1])
            throw new InvalidOperationException("Combo box was not created with the expected default selection.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComboBoxFormField.docx");
        doc.Save(outputPath);
    }
}

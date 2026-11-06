using System;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    // Reusable method to add a combo box form field.
    public static void AddComboBoxFormField(Document doc, string fieldName, List<string> items, int defaultIndex)
    {
        if (doc == null) throw new ArgumentNullException(nameof(doc));
        if (string.IsNullOrEmpty(fieldName)) throw new ArgumentException("Field name must be provided.", nameof(fieldName));
        if (items == null) throw new ArgumentNullException(nameof(items));
        if (items.Count == 0) throw new ArgumentException("At least one item must be supplied.", nameof(items));
        if (defaultIndex < 0 || defaultIndex >= items.Count)
            throw new ArgumentOutOfRangeException(nameof(defaultIndex), "Default index is out of range.");

        // Use DocumentBuilder to insert the combo box.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertComboBox(fieldName, items.ToArray(), defaultIndex);
    }

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Define combo box items.
        List<string> comboItems = new List<string> { "Red", "Green", "Blue", "Yellow" };
        int defaultItemIndex = 2; // "Blue"

        // Add the combo box form field.
        AddComboBoxFormField(doc, "ColorChoice", comboItems, defaultItemIndex);

        // Validate that the field was added.
        var comboField = doc.Range.FormFields["ColorChoice"];
        if (comboField == null)
            throw new InvalidOperationException("Combo box form field was not created.");

        // Output the default selected value to the console.
        Console.WriteLine($"Combo box '{comboField.Name}' default selected value: {comboField.Result}");

        // Save the document.
        doc.Save("ComboBoxFormField.docx");
    }
}

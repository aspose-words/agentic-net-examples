using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace ComboBoxFormFieldExample
{
    public class Program
    {
        // Reusable method that adds a combo box form field to the specified document.
        // Parameters:
        //   doc          - The Aspose.Words Document to modify.
        //   fieldName    - Unique name for the combo box form field.
        //   items        - Array of string items to populate the combo box.
        //   defaultIndex - Zero‑based index of the item that should be selected by default.
        public static void AddComboBox(Document doc, string fieldName, string[] items, int defaultIndex)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (string.IsNullOrEmpty(fieldName)) throw new ArgumentException("Field name must be provided.", nameof(fieldName));
            if (items == null || items.Length == 0) throw new ArgumentException("At least one item must be supplied.", nameof(items));
            if (defaultIndex < 0 || defaultIndex >= items.Length) throw new ArgumentOutOfRangeException(nameof(defaultIndex));

            // Use DocumentBuilder to insert the combo box at the end of the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertComboBox(fieldName, items, defaultIndex);

            // Validate that the field was added correctly.
            FormField comboBox = doc.Range.FormFields[fieldName];
            if (comboBox == null)
                throw new InvalidOperationException($"Form field '{fieldName}' was not found after insertion.");

            // Ensure the field type is ComboBox.
            if (comboBox.Type != FieldType.FieldFormDropDown)
                throw new InvalidOperationException($"Form field '{fieldName}' is not a combo box.");

            // Verify the default selected value matches the expected item.
            string expectedValue = items[defaultIndex];
            if (!string.Equals(comboBox.Result, expectedValue, StringComparison.Ordinal))
                throw new InvalidOperationException($"Default value mismatch. Expected '{expectedValue}', got '{comboBox.Result}'.");
        }

        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Define combo box parameters.
            string comboBoxName = "CountrySelector";
            string[] countryItems = new[] { "USA", "Canada", "Mexico", "Germany", "France" };
            int defaultItemIndex = 2; // Select "Mexico" by default.

            // Add the combo box form field.
            AddComboBox(doc, comboBoxName, countryItems, defaultItemIndex);

            // Save the document to disk.
            string outputPath = "ComboBoxFormField.docx";
            doc.Save(outputPath);
        }
    }
}

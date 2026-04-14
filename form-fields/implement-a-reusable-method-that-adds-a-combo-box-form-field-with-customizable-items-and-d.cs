using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

namespace ComboBoxFormFieldExample
{
    public class Program
    {
        /// <summary>
        /// Inserts a combo box form field with the specified name, items and default selected index.
        /// Validates the insertion and the default result.
        /// </summary>
        private static FormField InsertComboBoxField(DocumentBuilder builder, string name, string[] items, int defaultIndex)
        {
            if (builder == null) throw new ArgumentNullException(nameof(builder));
            if (string.IsNullOrEmpty(name)) throw new ArgumentException("Form field name cannot be null or empty.", nameof(name));
            if (items == null || items.Length == 0) throw new ArgumentException("Items collection cannot be null or empty.", nameof(items));
            if (defaultIndex < 0 || defaultIndex >= items.Length) throw new ArgumentOutOfRangeException(nameof(defaultIndex), "Default index must be within the items range.");

            // Insert the combo box using the prescribed API.
            FormField comboBox = builder.InsertComboBox(name, items, defaultIndex);

            // Ensure the field was created.
            if (comboBox == null)
                throw new InvalidOperationException("Failed to insert the combo box form field.");

            // Verify that the result matches the expected default item.
            if (!comboBox.Result.Equals(items[defaultIndex], StringComparison.Ordinal))
                throw new InvalidOperationException("The combo box result does not match the default selected item.");

            return comboBox;
        }

        public static void Main()
        {
            // Create a new document and a builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some introductory text.
            builder.Writeln("Please select a fruit from the list below:");

            // Define the items for the combo box.
            string[] fruits = { "Apple", "Banana", "Cherry" };
            int defaultIndex = 1; // Banana will be selected by default.

            // Insert the combo box using the reusable method.
            FormField fruitCombo = InsertComboBoxField(builder, "FruitCombo", fruits, defaultIndex);

            // Update the selected index to demonstrate reading and writing.
            fruitCombo.DropDownSelectedIndex = 2; // Change selection to "Cherry".

            // Validate the update.
            if (fruitCombo.DropDownSelectedIndex != 2)
                throw new InvalidOperationException("Failed to update the selected index of the combo box.");

            if (!fruitCombo.Result.Equals(fruits[2], StringComparison.Ordinal))
                throw new InvalidOperationException("The combo box result does not reflect the updated selection.");

            // Ensure at least one form field exists before saving.
            if (doc.Range.FormFields.Count == 0)
                throw new InvalidOperationException("The document does not contain any form fields.");

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComboBoxFormField.docx");
            doc.Save(outputPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}

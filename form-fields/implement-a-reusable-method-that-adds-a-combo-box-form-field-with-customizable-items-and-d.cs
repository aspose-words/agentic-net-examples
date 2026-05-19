using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsFormFieldsDemo
{
    public class Program
    {
        // Reusable method that inserts a combo box form field.
        // Parameters:
        //   builder      - DocumentBuilder positioned where the field should be inserted.
        //   name         - Name of the form field (also creates a bookmark with the same name).
        //   items        - Array of strings representing the items in the combo box (max 25).
        //   defaultIndex - Zero‑based index of the item that will be selected by default.
        // Returns the inserted FormField instance.
        public static FormField AddComboBox(DocumentBuilder builder, string name, string[] items, int defaultIndex)
        {
            if (builder == null) throw new ArgumentNullException(nameof(builder));
            if (string.IsNullOrEmpty(name)) throw new ArgumentException("Form field name cannot be null or empty.", nameof(name));
            if (items == null || items.Length == 0) throw new ArgumentException("Items collection cannot be null or empty.", nameof(items));
            if (defaultIndex < 0 || defaultIndex >= items.Length) throw new ArgumentOutOfRangeException(nameof(defaultIndex));

            // Insert the combo box using the Aspose.Words API.
            FormField comboBox = builder.InsertComboBox(name, items, defaultIndex);

            // Validate that the field was created.
            if (comboBox == null)
                throw new InvalidOperationException("Failed to insert the combo box form field.");

            // Additional validation: ensure the selected index matches the requested default.
            if (comboBox.DropDownSelectedIndex != defaultIndex)
                throw new InvalidOperationException("The combo box default index was not set correctly.");

            return comboBox;
        }

        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write a prompt before the combo box.
            builder.Write("Select your favorite fruit: ");

            // Define items for the combo box.
            string[] fruitItems = { "Apple", "Banana", "Cherry", "Date" };
            int defaultFruitIndex = 1; // Banana will be selected by default.

            // Use the reusable method to add the combo box.
            FormField fruitComboBox = AddComboBox(builder, "FruitCombo", fruitItems, defaultFruitIndex);

            // Verify that the field exists in the document's form field collection.
            FormField retrievedField = doc.Range.FormFields["FruitCombo"];
            if (retrievedField == null)
                throw new InvalidOperationException("The combo box form field was not found in the document.");

            // Optionally, change the selected item after creation.
            retrievedField.DropDownSelectedIndex = 2; // Select "Cherry".
            // The Result property reflects the currently selected item.
            Console.WriteLine($"Combo box result after change: {retrievedField.Result}");

            // Save the document to the output folder.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComboBoxFormField.docx");
            doc.Save(outputPath);
        }
    }
}

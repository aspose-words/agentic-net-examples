using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace ComboBoxFormFieldExample
{
    public class Program
    {
        /// <summary>
        /// Adds a combo box (drop‑down) form field to the document at the builder's current position.
        /// </summary>
        /// <param name="builder">The DocumentBuilder positioned where the field should be inserted.</param>
        /// <param name="name">The name of the form field (also creates a bookmark with the same name).</param>
        /// <param name="items">Array of strings that will appear as selectable items.</param>
        /// <param name="defaultIndex">Zero‑based index of the item that should be selected by default.</param>
        /// <returns>The inserted FormField instance.</returns>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when defaultIndex is outside the items array.</exception>
        /// <exception cref="InvalidOperationException">Thrown when the field cannot be retrieved after insertion.</exception>
        public static FormField AddComboBoxFormField(DocumentBuilder builder, string name, string[] items, int defaultIndex)
        {
            if (items == null) throw new ArgumentNullException(nameof(items));
            if (defaultIndex < 0 || defaultIndex >= items.Length)
                throw new ArgumentOutOfRangeException(nameof(defaultIndex), "Default index must be within the items array.");

            // Insert the combo box using the Aspose.Words API.
            FormField comboBox = builder.InsertComboBox(name, items, defaultIndex);

            // Verify that the field was added to the document.
            FormField retrieved = builder.Document.Range.FormFields[name];
            if (retrieved == null)
                throw new InvalidOperationException($"Form field '{name}' was not found after insertion.");

            // Optionally, you could adjust additional properties here (e.g., Enable, HelpText, etc.).

            return comboBox;
        }

        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write a prompt before the combo box.
            builder.Writeln("Select your favorite color:");

            // Define items for the combo box.
            string[] colors = { "Red", "Green", "Blue", "Yellow" };
            int defaultSelection = 2; // "Blue" will be selected by default.

            // Add the combo box using the reusable method.
            FormField colorField = AddComboBoxFormField(builder, "FavoriteColor", colors, defaultSelection);

            // Output the selected value to the console (no user interaction required).
            Console.WriteLine($"Combo box '{colorField.Name}' inserted with default selection: {colorField.Result}");

            // Save the document to the file system.
            string outputPath = "ComboBoxFormField.docx";
            doc.Save(outputPath);
        }
    }
}

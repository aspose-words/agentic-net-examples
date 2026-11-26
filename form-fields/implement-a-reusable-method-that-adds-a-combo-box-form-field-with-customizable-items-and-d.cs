using System;
using Aspose.Words;

namespace FormFieldExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Define combo box parameters.
            string comboBoxName = "CountryComboBox";
            string[] items = { "USA", "Canada", "Mexico", "Germany", "France" };
            int defaultIndex = 2; // Select "Mexico" by default.

            // Add the combo box to the document.
            AddComboBox(doc, comboBoxName, items, defaultIndex);

            // Save the document.
            doc.Save("ComboBoxFormField.docx");
        }

        /// <summary>
        /// Adds a combo box (dropdown) form field to the specified document.
        /// </summary>
        /// <param name="doc">The document to modify.</param>
        /// <param name="fieldName">The name of the form field.</param>
        /// <param name="items">Array of items to display in the dropdown.</param>
        /// <param name="defaultIndex">Zero‑based index of the item that should be selected by default.</param>
        static void AddComboBox(Document doc, string fieldName, string[] items, int defaultIndex)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (string.IsNullOrEmpty(fieldName)) throw new ArgumentException("Field name must be provided.", nameof(fieldName));
            if (items == null || items.Length == 0) throw new ArgumentException("At least one item must be supplied.", nameof(items));
            if (defaultIndex < 0 || defaultIndex >= items.Length) throw new ArgumentOutOfRangeException(nameof(defaultIndex), "Default index is out of range.");

            // Use DocumentBuilder to insert the combo box at the current cursor position.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertComboBox(fieldName, items, defaultIndex);
        }
    }
}

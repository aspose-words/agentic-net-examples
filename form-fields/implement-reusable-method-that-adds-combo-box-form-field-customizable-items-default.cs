using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace FormFieldUtilities
{
    /// <summary>
    /// Helper methods for working with form fields.
    /// </summary>
    public static class FormFieldHelper
    {
        /// <summary>
        /// Inserts a combo box (drop‑down) form field at the current builder position.
        /// </summary>
        /// <param name="builder">The DocumentBuilder that is positioned where the field should be inserted.</param>
        /// <param name="name">The name of the form field (bookmark will be created automatically). Empty string is allowed.</param>
        /// <param name="items">Array of strings that will appear as selectable items. Maximum 25 items as per Word limits.</param>
        /// <param name="defaultIndex">
        /// Zero‑based index of the item that should be selected by default.
        /// Use -1 to leave the field empty (no selection).
        /// </param>
        /// <returns>The inserted FormField instance.</returns>
        public static FormField AddComboBox(DocumentBuilder builder, string name, string[] items, int defaultIndex)
        {
            // Aspose.Words provides InsertComboBox which creates all required field nodes.
            // The method returns the FormField that represents the newly inserted combo box.
            FormField comboBox = builder.InsertComboBox(name, items, defaultIndex);

            // Optionally you can adjust additional properties here, e.g. enable calculation on exit.
            // comboBox.CalculateOnExit = true;

            return comboBox;
        }
    }

    /// <summary>
    /// Demonstrates the reusable method.
    /// </summary>
    class Program
    {
        static void Main()
        {
            // Create a new document and a builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write a prompt before the combo box.
            builder.Write("Pick a fruit: ");

            // Define items and default selection.
            string[] fruitItems = { "Apple", "Banana", "Cherry", "Date" };
            int defaultSelected = 1; // "Banana" will be selected initially.

            // Use the reusable helper to insert the combo box.
            FormField fruitCombo = FormFieldHelper.AddComboBox(builder, "FruitCombo", fruitItems, defaultSelected);

            // Verify that the field was created correctly (optional).
            Console.WriteLine($"Combo box name: {fruitCombo.Name}");
            Console.WriteLine($"Selected index: {fruitCombo.DropDownSelectedIndex}");
            Console.WriteLine($"Result text: {fruitCombo.Result}");

            // Save the document.
            doc.Save("ComboBoxFormField.docx");
        }
    }
}

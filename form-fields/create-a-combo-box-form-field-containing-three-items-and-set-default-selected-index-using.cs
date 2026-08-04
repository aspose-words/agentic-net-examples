using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsFormFieldsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a prompt before the combo box.
            builder.Write("Pick a fruit: ");

            // Define the items that will appear in the combo box.
            string[] items = { "Apple", "Banana", "Cherry" };

            // Insert a combo box form field named "FruitCombo" with the items above.
            // The third parameter (selectedIndex) sets the default selected item (0‑based).
            // Here we set it to 1, so "Banana" will be selected by default.
            FormField comboBox = builder.InsertComboBox("FruitCombo", items, 1);

            // Save the document to a file.
            doc.Save("ComboBoxFormField.docx");
        }
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsFormFieldsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a prompt before the combo box.
            builder.Write("Pick a fruit: ");

            // Define the items that will appear in the combo box.
            string[] items = { "Apple", "Banana", "Cherry" };

            // Insert the combo box form field with a name, the items, and the default selected index (0 = "Apple").
            FormField comboBox = builder.InsertComboBox("FruitCombo", items, 0);

            // Save the document to disk.
            doc.Save("ComboBoxFormField.docx");
        }
    }
}

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

            // Write a prompt before the combo box.
            builder.Write("Pick a fruit: ");

            // Define the items that will appear in the combo box.
            string[] items = { "Apple", "Banana", "Cherry" };

            // Insert a combo box form field with a name, the items, and set the default selected index.
            // In this example the second item ("Banana") is selected by default (index 1).
            FormField comboBox = builder.InsertComboBox("FruitCombo", items, 1);

            // Save the document to the file system.
            doc.Save("ComboBoxFormField.docx");
        }
    }
}

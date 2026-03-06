using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsComboBoxExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write some prompt text before the combo box.
            builder.Write("Pick a fruit: ");

            // Define the items that will appear in the combo box.
            string[] items = { "Apple", "Banana", "Cherry" };

            // Insert the combo box form field.
            // Parameters: name of the field, array of items, index of the initially selected item.
            builder.InsertComboBox("FruitCombo", items, 0);

            // Save the document to a .docx file.
            doc.Save("ComboBox.docx");
        }
    }
}

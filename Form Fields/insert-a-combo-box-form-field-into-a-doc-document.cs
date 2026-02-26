using System;
using Aspose.Words;
using Aspose.Words.Fields;

class InsertComboBoxExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some prompt text before the combo box.
        builder.Write("Pick a fruit: ");

        // Define the items that will appear in the combo box.
        string[] items = { "Apple", "Banana", "Cherry" };

        // Insert the combo box form field.
        // Parameters: name of the field, array of items, index of the initially selected item.
        builder.InsertComboBox("FruitComboBox", items, 0);

        // Save the document to disk.
        doc.Save("ComboBox.docx");
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Fields;

class InsertComboBoxExample
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("InputDocument.docx");

        // Create a DocumentBuilder to modify the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some prompt text before the combo box.
        builder.Write("Please select a fruit: ");

        // Define the items that will appear in the combo box.
        string[] items = { "Apple", "Banana", "Cherry" };

        // Insert the combo box form field at the current cursor position.
        // Parameters: name of the field, array of items, index of the initially selected item.
        builder.InsertComboBox("FruitComboBox", items, 0);

        // Save the modified document.
        doc.Save("OutputDocument.docx");
    }
}

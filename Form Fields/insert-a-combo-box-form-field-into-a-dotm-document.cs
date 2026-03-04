using System;
using Aspose.Words;
using Aspose.Words.Fields;

class InsertComboBoxIntoDotm
{
    static void Main()
    {
        // Load an existing DOTM template (or create a new document if you prefer).
        Document doc = new Document("Template.dotm");

        // Create a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some prompt text before the combo box.
        builder.Write("Pick a fruit: ");

        // Define the items that will appear in the combo box.
        string[] items = { "Apple", "Banana", "Cherry" };

        // Insert the combo box form field at the current cursor position.
        // Parameters: field name, items array, index of the initially selected item.
        builder.InsertComboBox("FruitCombo", items, 0);

        // Save the modified document as a DOTM file.
        doc.Save("Result.dotm");
    }
}

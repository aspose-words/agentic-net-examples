using System;
using Aspose.Words;
using Aspose.Words.Fields;

class InsertComboBoxIntoDotx
{
    static void Main()
    {
        // Load an existing DOTX template.
        Document doc = new Document("Template.dotx");

        // Create a DocumentBuilder for the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define the items that will appear in the combo box.
        string[] comboItems = { "Option 1", "Option 2", "Option 3" };

        // Insert the combo box form field.
        // Parameters: field name, items array, index of the default selected item.
        builder.InsertComboBox("MyComboBox", comboItems, 0);

        // Save the document back as a DOTX file.
        doc.Save("Result.dotx");
    }
}

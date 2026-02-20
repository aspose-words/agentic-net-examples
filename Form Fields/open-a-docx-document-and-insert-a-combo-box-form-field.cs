using System;
using Aspose.Words;
using Aspose.Words.Fields;

class InsertComboBoxFormField
{
    static void Main()
    {
        // Load an existing DOCX document.
        // Replace "input.docx" with the path to your source file.
        Document doc = new Document("input.docx");

        // Create a DocumentBuilder attached to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define the items that will appear in the combo box.
        string[] comboItems = { "Option 1", "Option 2", "Option 3" };

        // Insert a combo box form field at the current cursor position.
        // Parameters: field name, array of items, index of the default selected item.
        builder.InsertComboBox("MyComboBox", comboItems, 0);

        // Save the modified document.
        // Replace "output.docx" with the desired output path.
        doc.Save("output.docx");
    }
}

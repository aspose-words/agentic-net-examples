using System;
using Aspose.Words;
using Aspose.Words.Fields;

class InsertComboBoxIntoDotx
{
    static void Main()
    {
        // Load the DOTX template.
        Document doc = new Document("Template.dotx");

        // Create a DocumentBuilder for the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a prompt before the combo box.
        builder.Write("Pick a fruit: ");

        // Define the items for the combo box.
        string[] items = { "Apple", "Banana", "Cherry" };

        // Insert the combo box form field at the current cursor position.
        // Parameters: field name, items array, selected index (0 = first item).
        FormField comboBox = builder.InsertComboBox("FruitComboBox", items, 0);

        // Optionally, set additional properties (e.g., enable the field).
        comboBox.Enabled = true;

        // Save the modified document. You can save as DOTX to keep it as a template,
        // or as DOCX if you want a regular document.
        doc.Save("TemplateWithComboBox.dotx");
    }
}

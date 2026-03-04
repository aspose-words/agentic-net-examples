using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a prompt before the combo box.
        builder.Write("Pick a fruit: ");

        // Define the items that will appear in the combo box.
        string[] items = { "Apple", "Banana", "Cherry" };

        // Insert the combo box form field.
        // Parameters: name of the field, array of items, index of the initially selected item.
        FormField comboBox = builder.InsertComboBox("FruitCombo", items, 0);

        // Optionally, you can access the combo box properties, e.g.:
        // comboBox.DropDownSelectedIndex = 1; // Select "Banana" by default.

        // Save the document to a file.
        doc.Save("ComboBoxFormField.docx");
    }
}

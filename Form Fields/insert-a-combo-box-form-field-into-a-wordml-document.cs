using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder which will be used to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some introductory text before the combo box.
        builder.Write("Pick a fruit: ");

        // Define the list of items that will appear in the combo box.
        string[] items = { "Apple", "Banana", "Cherry" };

        // Insert the combo box form field at the current cursor position.
        // Parameters: field name, array of items, index of the initially selected item.
        FormField comboBox = builder.InsertComboBox("FruitCombo", items, 0);

        // Optional: make the field recalculate any dependent fields when the user exits it.
        comboBox.CalculateOnExit = true;

        // Save the document to a file.
        doc.Save("ComboBoxFormField.docx");
    }
}

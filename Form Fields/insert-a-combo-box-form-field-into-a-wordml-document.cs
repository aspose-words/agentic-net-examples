using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a prompt before the combo box.
        builder.Write("Pick a fruit: ");

        // Define the items that will appear in the combo box.
        string[] items = { "Apple", "Banana", "Cherry" };

        // Insert the combo box form field.
        // Parameters: field name, array of items, index of the initially selected item.
        FormField comboBox = builder.InsertComboBox("FruitCombo", items, 0);

        // Optionally set the displayed result (the selected item's text).
        comboBox.Result = items[0];

        // Save the document to a DOCX file.
        doc.Save("ComboBoxFormField.docx");
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Fields;

class InsertComboBoxIntoDot
{
    static void Main()
    {
        // Create a new blank document. This will be saved as a DOT (Word template) later.
        Document doc = new Document();

        // Initialize a DocumentBuilder which allows us to add content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some introductory text before the combo box.
        builder.Write("Please select a fruit: ");

        // Define the items that will appear in the combo box.
        string[] items = { "Apple", "Banana", "Cherry" };

        // Insert the combo box form field.
        // Parameters: field name, array of items, index of the initially selected item.
        FormField comboBox = builder.InsertComboBox("FruitComboBox", items, 0);

        // Optionally, you can modify properties of the inserted form field.
        // For example, ensure the field is enabled and calculate its value on exit.
        comboBox.Enabled = true;
        comboBox.CalculateOnExit = true;

        // Save the document as a DOT file (Word template).
        // The SaveFormat.Dot constant ensures the correct file type.
        doc.Save("FruitTemplate.dot", SaveFormat.Dot);
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("InputDocument.docx");

        // Create a DocumentBuilder to modify the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some prompt text before the combo box.
        builder.Write("Please select a fruit: ");

        // Define the items for the combo box.
        string[] items = { "Apple", "Banana", "Cherry" };

        // Insert the combo box form field at the current cursor position.
        // Parameters: name of the field, array of items, index of the initially selected item.
        FormField comboBox = builder.InsertComboBox("FruitComboBox", items, 0);

        // Optionally, you can set additional properties on the form field.
        // For example, make the field enabled (default) and set a help text.
        comboBox.Enabled = true;
        comboBox.HelpText = "Choose your favorite fruit.";

        // Save the modified document.
        doc.Save("OutputDocument.docx");
    }
}

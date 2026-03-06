using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOTM template or create a new empty document.
        // If you have a template file, replace the constructor argument with its path.
        Document doc = new Document(); // new empty document
        // Document doc = new Document("ExistingTemplate.dotm");

        // Create a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some prompt text before the combo box.
        builder.Write("Pick a fruit: ");

        // Define the items that will appear in the combo box.
        string[] items = { "Apple", "Banana", "Cherry" };

        // Insert the combo box form field.
        // Parameters: field name, array of items, index of the initially selected item.
        FormField comboBox = builder.InsertComboBox("FruitCombo", items, 0);

        // (Optional) Set additional properties, e.g., recalculate on exit.
        // comboBox.CalculateOnExit = true;

        // Save the document as a macro‑enabled template (.dotm).
        doc.Save("ComboBoxTemplate.dotm", SaveFormat.Dotm);
    }
}

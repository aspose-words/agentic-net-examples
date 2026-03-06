using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some introductory text.
        builder.Write("Pick a fruit: ");

        // Items that will appear in the combo box.
        string[] items = { "Apple", "Banana", "Cherry" };

        // Insert the combo box form field. The first item (index 0) is selected by default.
        builder.InsertComboBox("FruitCombo", items, 0);

        // Save the document as a plain‑text file.
        doc.Save("ComboBox.txt", SaveFormat.Text);
    }
}

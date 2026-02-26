using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new document. The file will be saved as a DOTM (macro‑enabled template).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a prompt for the user.
        builder.Writeln("Pick a fruit:");

        // Items that will appear in the combo box.
        string[] items = { "Apple", "Banana", "Cherry" };

        // Insert the combo box form field at the current cursor position.
        // The first item ("Apple") will be selected by default (selectedIndex = 0).
        builder.InsertComboBox("FruitCombo", items, 0);

        // Save the document as a DOTM file.
        doc.Save("ComboBoxTemplate.dotm");
    }
}

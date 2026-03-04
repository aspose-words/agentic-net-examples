using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some explanatory text before the combo box.
        builder.Write("Pick a fruit: ");

        // Define the items that will appear in the combo box.
        string[] items = { "Apple", "Banana", "Cherry" };

        // Insert the combo box form field at the current cursor position.
        // The third argument (0) selects the first item ("Apple") by default.
        builder.InsertComboBox("FruitCombo", items, 0);

        // Save the document as a plain‑text file.
        doc.Save("ComboBoxDocument.txt", SaveFormat.Text);
    }
}

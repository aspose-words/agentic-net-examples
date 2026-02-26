using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document (DOT template)
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some introductory text
        builder.Write("Select a fruit: ");

        // Define the items that will appear in the combo box
        string[] items = { "Apple", "Banana", "Cherry" };

        // Insert the combo box form field at the current cursor position.
        // The first item ("Apple") will be selected by default (index 0).
        builder.InsertComboBox("FruitCombo", items, 0);

        // Save the document as a Word template (.dot)
        doc.Save("ComboBoxTemplate.dot");
    }
}

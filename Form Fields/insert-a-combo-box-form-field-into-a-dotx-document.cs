using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document which will be saved as a DOTX template.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some introductory text before the combo box.
        builder.Write("Select a fruit: ");

        // Define the items that will appear in the combo box.
        string[] items = { "Apple", "Banana", "Cherry" };

        // Insert the combo box form field.
        // Parameters: field name, array of items, index of the initially selected item.
        builder.InsertComboBox("FruitCombo", items, 0);

        // Save the document as a DOTX template.
        doc.Save("ComboBoxTemplate.dotx", SaveFormat.Dotx);
    }
}

using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new document (or load an existing DOCM with new Document("input.docm"))
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some introductory text
        builder.Write("Pick a fruit: ");

        // Items for the combo box
        string[] items = { "Apple", "Banana", "Cherry" };

        // Insert the combo box form field; the first item (index 0) is selected by default
        FormField comboBox = builder.InsertComboBox("FruitCombo", items, 0);

        // Example of setting an additional property
        comboBox.CalculateOnExit = true;

        // Save the document as a DOCM file
        doc.Save("ComboBoxFormField.docm");
    }
}

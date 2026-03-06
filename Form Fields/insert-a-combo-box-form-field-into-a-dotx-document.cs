using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the DOTX template.
        Document doc = new Document("Template.dotx");

        // Initialize a DocumentBuilder for the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some prompt text before the combo box.
        builder.Write("Pick a fruit: ");

        // Define the items that will appear in the combo box.
        string[] items = { "Apple", "Banana", "Cherry" };

        // Insert the combo box form field at the current cursor position.
        // The first item ("Apple") will be selected by default (index 0).
        builder.InsertComboBox("FruitCombo", items, 0);

        // Save the modified document.
        doc.Save("Result.docx");
    }
}

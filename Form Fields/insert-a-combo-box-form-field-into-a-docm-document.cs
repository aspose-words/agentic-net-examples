using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some prompt text before the combo box.
        builder.Write("Please select a fruit: ");

        // Define the items that will appear in the combo box.
        string[] items = { "Apple", "Banana", "Cherry" };

        // Insert the combo box form field.
        // Parameters: name of the field, array of items, index of the initially selected item.
        FormField comboBox = builder.InsertComboBox("FruitComboBox", items, 0);

        // Optionally, you can modify properties of the inserted form field.
        // For example, set the field to be enabled and calculate its value on exit.
        comboBox.Enabled = true;
        comboBox.CalculateOnExit = true;

        // Save the document as a DOCM (macro-enabled) file.
        // The file extension determines the format.
        doc.Save("ComboBoxFormField.docm");
    }
}

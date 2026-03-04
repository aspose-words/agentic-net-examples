using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some explanatory text.
        builder.Write("Pick a fruit: ");

        // Define the items that will appear in the combo box.
        string[] items = { "Apple", "Banana", "Cherry" };

        // Insert the combo box form field. The first item (index 0) is selected by default.
        builder.InsertComboBox("FruitCombo", items, 0);

        // Prepare save options for MHTML. Keep the combo box as an interactive <select> element.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml);
        saveOptions.ExportDropDownFormFieldAsText = false; // preserve dropdown functionality

        // Save the document as MHTML.
        doc.Save("ComboBox.mhtml", saveOptions);
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertComboBoxIntoMhtml
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some introductory text.
        builder.Write("Please select a fruit: ");

        // Define the items for the combo box.
        string[] items = { "Apple", "Banana", "Cherry" };

        // Insert the combo box form field. The third argument (0) selects the first item by default.
        builder.InsertComboBox("FruitComboBox", items, 0);

        // Configure save options for MHTML.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml);
        // Keep the combo box as a selectable <select> element in the output.
        saveOptions.ExportDropDownFormFieldAsText = false;

        // Save the document as MHTML.
        doc.Save("ComboBoxDocument.mhtml", saveOptions);
    }
}

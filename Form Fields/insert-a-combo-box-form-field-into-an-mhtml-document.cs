using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;

class InsertComboBoxIntoMhtml
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some prompt text before the combo box.
        builder.Write("Pick a fruit: ");

        // Define the items that will appear in the combo box.
        string[] items = { "Apple", "Banana", "Cherry" };

        // Insert the combo box form field.
        // Parameters: name of the field, array of items, index of the initially selected item.
        builder.InsertComboBox("FruitComboBox", items, 0);

        // Configure HTML/MHTML save options.
        // ExportDropDownFormFieldAsText = false ensures the combo box is saved as a <select> element.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportDropDownFormFieldAsText = false,
            // Save as MHTML (MIME HTML) format.
            SaveFormat = SaveFormat.Mhtml
        };

        // Save the document as an MHTML file.
        doc.Save("ComboBoxDocument.mhtml", saveOptions);
    }
}

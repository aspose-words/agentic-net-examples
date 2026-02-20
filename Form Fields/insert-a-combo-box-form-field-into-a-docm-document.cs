using System;
using Aspose.Words;
using Aspose.Words.Fields;

class InsertComboBoxFormField
{
    static void Main()
    {
        // Create a new, empty document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document for editing.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Optional: add a paragraph before the form field.
        builder.InsertParagraph();

        // Define the items that will appear in the combo box.
        string[] comboItems = new[]
        {
            "-- Select an option --",
            "Option A",
            "Option B",
            "Option C"
        };

        // Insert the combo box form field.
        // Parameters: field name, array of items, index of the default selected item.
        builder.InsertComboBox("MyComboBox", comboItems, 0);

        // Save the document as a macro‑enabled DOCM file.
        doc.Save("ComboBoxFormField.docm", SaveFormat.Docm);
    }
}

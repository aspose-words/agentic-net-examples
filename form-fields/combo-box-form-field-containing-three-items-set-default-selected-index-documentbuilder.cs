using System;
using Aspose.Words;
using Aspose.Words.Fields;

class ComboBoxExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Optional: write some prompt text before the combo box.
        builder.Write("Select an option: ");

        // Define the items that will appear in the combo box.
        string[] items = { "First", "Second", "Third" };

        // Insert the combo box form field.
        // Parameters: field name, items array, default selected index (0‑based).
        // Here we set the default selected index to 1, so "Second" is selected by default.
        builder.InsertComboBox("MyComboBox", items, 1);

        // Save the document to disk.
        doc.Save("ComboBoxFormField.docx");
    }
}

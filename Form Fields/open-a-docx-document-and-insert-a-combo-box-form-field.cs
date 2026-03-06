using System;
using Aspose.Words;
using Aspose.Words.Fields;

class InsertComboBoxExample
{
    static void Main()
    {
        // Load an existing DOCX document from disk.
        // This uses the Document(string) constructor, which is the standard load operation.
        Document doc = new Document("input.docx");

        // Create a DocumentBuilder for the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some prompt text before the combo box (optional).
        builder.Write("Please select a fruit: ");

        // Define the items that will appear in the combo box.
        string[] items = { "Apple", "Banana", "Cherry" };

        // Insert the combo box form field at the current cursor position.
        // Parameters: field name, items array, index of the initially selected item.
        FormField comboBox = builder.InsertComboBox("MyComboBox", items, 0);

        // (Optional) Verify that the field was inserted correctly.
        // Console.WriteLine($"ComboBox Name: {comboBox.Name}, Selected: {comboBox.Result}");

        // Save the modified document to a new file.
        // This uses the Document.Save method, which is the standard save operation.
        doc.Save("output.docx");
    }
}

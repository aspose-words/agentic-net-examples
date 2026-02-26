using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document.
        // Replace "input.docx" with the path to your source file.
        Document doc = new Document("input.docx");

        // Create a DocumentBuilder for the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some prompt text before the combo box (optional).
        builder.Write("Please select a fruit: ");

        // Define the items that will appear in the combo box.
        string[] items = { "Apple", "Banana", "Cherry" };

        // Insert the combo box form field at the current cursor position.
        // Parameters: field name, items array, selected index (0‑based).
        FormField comboBox = builder.InsertComboBox("FruitComboBox", items, 0);

        // Optionally, you can verify properties of the inserted field.
        // Console.WriteLine($"ComboBox Name: {comboBox.Name}");
        // Console.WriteLine($"Default Selected Item: {comboBox.Result}");

        // Save the modified document.
        // Replace "output.docx" with the desired output path.
        doc.Save("output.docx");
    }
}

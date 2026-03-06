using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the existing DOCX file to load.
        string inputPath = "InputDocument.docx";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Create a DocumentBuilder for editing the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a prompt before the combo box.
        builder.Write("Pick a fruit: ");

        // Define the items that will appear in the combo box.
        string[] items = { "Apple", "Banana", "Cherry" };

        // Insert the combo box form field at the current cursor position.
        // Parameters: field name, items array, selected index (0 = first item).
        builder.InsertComboBox("FruitCombo", items, 0);

        // Path to save the modified document.
        string outputPath = "OutputDocument.docx";

        // Save the document with the newly inserted combo box.
        doc.Save(outputPath);
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Fields;

class InsertComboBoxFormField
{
    static void Main()
    {
        // Path to the existing DOCX file.
        string inputPath = @"C:\Docs\input.docx";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Create a DocumentBuilder attached to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define the items that will appear in the combo box.
        string[] comboItems = { "Option A", "Option B", "Option C" };

        // Insert a combo box form field at the current cursor position.
        // Parameters: field name, array of items, index of the default selected item.
        FormField comboBox = builder.InsertComboBox("MyComboBox", comboItems, 0);

        // (Optional) Set additional properties, e.g., calculate the field on exit.
        comboBox.CalculateOnExit = true;

        // Path where the modified document will be saved.
        string outputPath = @"C:\Docs\output.docx";

        // Save the document with the newly inserted combo box.
        doc.Save(outputPath);
    }
}

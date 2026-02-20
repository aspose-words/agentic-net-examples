using System;
using Aspose.Words;
using Aspose.Words.Fields;

class InsertComboBoxIntoDotm
{
    static void Main()
    {
        // Path to the existing DOTM template.
        string templatePath = @"C:\Docs\Template.dotm";

        // Load the DOTM document.
        Document doc = new Document(templatePath);

        // Create a DocumentBuilder for editing the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define the items that will appear in the combo box.
        string[] comboItems = { "Select an option", "Option A", "Option B", "Option C" };

        // Insert the combo box form field.
        // Parameters: field name, array of items, index of the default selected item.
        builder.InsertComboBox("MyComboBox", comboItems, 0);

        // Save the modified document (still as a DOTM to preserve the macro-enabled format).
        string outputPath = @"C:\Docs\ComboBoxInserted.dotm";
        doc.Save(outputPath);
    }
}

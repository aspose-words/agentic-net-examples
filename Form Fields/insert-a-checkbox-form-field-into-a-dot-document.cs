using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOT (Word template) file
        Document doc = new Document("Template.dot");

        // Create a DocumentBuilder for the loaded document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a checkbox form field at the current cursor position
        // Parameters: name, isChecked, size (0 = auto size)
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 0);

        // Optionally configure additional properties of the checkbox
        checkBox.IsCheckBoxExactSize = false;   // Use automatic size
        checkBox.HelpText = "Click to toggle";
        checkBox.OwnHelp = true;

        // Save the modified template back to a DOT file (or another format if desired)
        doc.Save("TemplateWithCheckBox.dot");
    }
}

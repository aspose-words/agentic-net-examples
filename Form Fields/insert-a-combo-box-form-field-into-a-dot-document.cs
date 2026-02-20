using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOT template.
        Document doc = new Document("Template.dot");

        // Create a DocumentBuilder for the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define the items that will appear in the combo box.
        string[] items = { "One", "Two", "Three" };

        // Insert a combo box form field.
        // Parameters: field name, array of items, index of the default selected item.
        FormField comboBox = builder.InsertComboBox("MyComboBox", items, 0);

        // Optional: make the field recalculate dependent fields when the user exits it.
        comboBox.CalculateOnExit = true;

        // Save the modified document back to DOT format.
        doc.Save("Result.dot");
    }
}

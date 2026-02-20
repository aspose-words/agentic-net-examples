using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define the items that will appear in the combo box.
        string[] items = {
            "-- Select your favorite footwear --",
            "Sneakers",
            "Oxfords",
            "Flip-flops",
            "Other"
        };

        // Insert a paragraph to separate the combo box from preceding content (optional).
        builder.InsertParagraph();

        // Insert the combo box form field.
        // Parameters: field name, array of items, index of the default selected item.
        FormField comboBox = builder.InsertComboBox("MyComboBox", items, 0);

        // Example of setting an additional property: recalculate references when the field is exited.
        comboBox.CalculateOnExit = true;

        // Save the document to disk.
        doc.Save("ComboBoxForm.docx");
    }
}

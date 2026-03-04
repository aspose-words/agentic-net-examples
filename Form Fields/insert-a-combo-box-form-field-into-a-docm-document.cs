using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace ComboBoxFormFieldExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document (DOCM supports macros and form fields)
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write a prompt before the combo box
            builder.Write("Pick a fruit: ");

            // Define the items that will appear in the combo box
            string[] items = { "Apple", "Banana", "Cherry" };

            // Insert the combo box form field.
            // Parameters: name of the field, array of items, index of the initially selected item (0‑based)
            FormField comboBox = builder.InsertComboBox("FruitComboBox", items, 0);

            // Optional: you can modify properties of the inserted form field here
            // e.g., comboBox.CalculateOnExit = true;

            // Save the document as a DOCM file (macro‑enabled Word document)
            doc.Save("ComboBoxFormField.docm", SaveFormat.Docm);
        }
    }
}

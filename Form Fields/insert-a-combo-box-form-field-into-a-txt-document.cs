using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace InsertComboBoxIntoTxt
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Define the items that will appear in the combo box.
            string[] comboItems = new string[]
            {
                "-- Select an option --",
                "Apple",
                "Banana",
                "Cherry"
            };

            // Insert a combo box form field named "FruitCombo" with the defined items.
            // The last parameter (0) sets the default selected index.
            builder.InsertComboBox("FruitCombo", comboItems, 0);

            // Save the document as plain text. TxtSaveOptions can be used to control the export.
            TxtSaveOptions txtOptions = new TxtSaveOptions();
            doc.Save("ComboBox.txt", txtOptions);
        }
    }
}

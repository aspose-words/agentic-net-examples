using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsFormFieldsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();

            // Initialize DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write a prompt before the combo box.
            builder.Write("Select a fruit: ");

            // Define the items for the combo box.
            string[] items = { "Apple", "Banana", "Cherry" };

            // Insert the combo box form field.
            // Parameters: name, items array, selected index (0‑based).
            FormField comboBox = builder.InsertComboBox("FruitComboBox", items, 1); // Default selects "Banana"

            // Optionally, you can verify the inserted field.
            Console.WriteLine($"ComboBox Name: {comboBox.Name}");
            Console.WriteLine($"Default Selected Item: {comboBox.Result}");

            // Save the document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "ComboBoxFormField.docx");
            doc.Save(outputPath);
        }
    }
}

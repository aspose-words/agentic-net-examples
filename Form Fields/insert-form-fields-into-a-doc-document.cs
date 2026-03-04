using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsFormFieldsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Create a DocumentBuilder which simplifies inserting content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // ------------------------------------------------------------
            // Insert a checkbox form field.
            // Parameters: name, isChecked (default state), size (points).
            // ------------------------------------------------------------
            builder.Write("Accept terms and conditions: ");
            FormField checkBox = builder.InsertCheckBox("AcceptTerms", false, 50);
            // Optional: set additional properties.
            checkBox.HelpText = "Check to accept the terms.";
            checkBox.OwnHelp = true;

            builder.Writeln(); // Move to next line.

            // ------------------------------------------------------------
            // Insert a combo box (drop‑down) form field.
            // Parameters: name, list of items, selected index.
            // ------------------------------------------------------------
            builder.Write("Select your favorite fruit: ");
            string[] fruitItems = { "Apple", "Banana", "Cherry", "Date" };
            FormField comboBox = builder.InsertComboBox("FruitChoice", fruitItems, 0);
            // Optional: make the field recalculate when the user changes selection.
            comboBox.CalculateOnExit = true;

            builder.Writeln(); // Move to next line.

            // ------------------------------------------------------------
            // Insert a text input form field.
            // Parameters: name, type, format, default text, max length.
            // ------------------------------------------------------------
            builder.Write("Enter your full name: ");
            FormField textInput = builder.InsertTextInput(
                "FullName",                     // field name
                TextFormFieldType.Regular,      // allows any text
                "",                             // no specific format
                "John Doe",                     // placeholder/default text
                0);                             // 0 = unlimited length

            // Optional: set a macro that runs when the field loses focus.
            textInput.ExitMacro = "OnExitFullName";

            // ------------------------------------------------------------
            // Update all fields so that their results are calculated.
            // This is important for fields like checkboxes that may have a result.
            // ------------------------------------------------------------
            doc.UpdateFields();

            // ------------------------------------------------------------
            // Save the document to disk.
            // ------------------------------------------------------------
            string outputPath = "FormFields.docx";
            doc.Save(outputPath);

            Console.WriteLine($"Document with form fields saved to: {outputPath}");
        }
    }
}

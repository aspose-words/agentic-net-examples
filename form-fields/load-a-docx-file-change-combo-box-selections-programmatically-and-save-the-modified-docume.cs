using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

        // --------------------------------------------------------------
        // Step 1: Create a DOCX file that contains a combo box form field.
        // --------------------------------------------------------------
        Document createDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(createDoc);

        // Add explanatory text.
        builder.Writeln("Please select a fruit from the list below:");

        // Insert a combo box named "FruitCombo".
        // The second argument is the list of items, the third argument is the zero‑based index of the default item.
        string[] fruitItems = { "Apple", "Banana", "Cherry" };
        builder.InsertComboBox("FruitCombo", fruitItems, 0);

        // Save the document that will be used as the input file.
        createDoc.Save(inputPath);

        // --------------------------------------------------------------
        // Step 2: Load the document and modify the combo box selection.
        // --------------------------------------------------------------
        Document loadDoc = new Document(inputPath);

        // Retrieve the combo box form field by its name.
        FormField comboField = loadDoc.Range.FormFields["FruitCombo"];
        if (comboField == null)
        {
            throw new InvalidOperationException("The combo box form field 'FruitCombo' was not found.");
        }

        // Verify that the field is a combo box (legacy drop‑down).
        if (comboField.Type != FieldType.FieldFormDropDown)
        {
            throw new InvalidOperationException("The form field 'FruitCombo' is not a combo box.");
        }

        // Change the selected value to "Banana". The value must exist in the items list.
        const string newSelection = "Banana";
        comboField.Result = newSelection;

        // Validate that the assignment succeeded.
        if (!string.Equals(comboField.Result, newSelection, StringComparison.Ordinal))
        {
            throw new InvalidOperationException($"Failed to set the combo box value to '{newSelection}'.");
        }

        // --------------------------------------------------------------
        // Step 3: Save the modified document.
        // --------------------------------------------------------------
        loadDoc.Save(outputPath);
    }
}

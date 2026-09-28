using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX with a combo box (dropdown) form field.
        Document createDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(createDoc);

        // Insert a combo box named "MyCombo" with three options, defaulting to the first one.
        string[] items = new string[] { "Option A", "Option B", "Option C" };
        builder.InsertComboBox("MyCombo", items, 0);

        // Save the initial document.
        const string inputPath = "input.docx";
        createDoc.Save(inputPath);

        // Step 2: Load the document we just created.
        Document loadDoc = new Document(inputPath);

        // Access the form fields collection.
        var formFields = loadDoc.Range.FormFields;

        // Retrieve the combo box by name.
        var comboField = formFields["MyCombo"];
        if (comboField == null)
        {
            throw new InvalidOperationException("Combo box 'MyCombo' was not found in the document.");
        }

        // Change the selected value programmatically.
        // Ensure the new value exists in the dropdown items.
        string newSelection = "Option C";
        bool exists = false;
        foreach (string item in comboField.DropDownItems)
        {
            if (item == newSelection)
            {
                exists = true;
                break;
            }
        }

        if (!exists)
        {
            throw new InvalidOperationException($"The value '{newSelection}' is not a valid option for the combo box.");
        }

        comboField.Result = newSelection; // Update the selection.

        // Step 3: Save the modified document.
        const string outputPath = "output.docx";
        loadDoc.Save(outputPath);
    }
}

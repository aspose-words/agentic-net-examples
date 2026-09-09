using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a prompt before the combo box.
        builder.Write("Select a color: ");

        // Define the items for the combo box.
        string[] items = { "Red", "Green", "Blue" };

        // Insert the combo box with the default selected index set to 1 (Green).
        FormField comboBox = builder.InsertComboBox("ColorCombo", items, 1);

        // Validate that the form field was added.
        FormFieldCollection fields = doc.Range.FormFields;
        if (fields == null || fields.Count == 0)
            throw new InvalidOperationException("No form fields were created.");

        // Retrieve the combo box by its name and verify the selected index.
        FormField retrieved = fields["ColorCombo"];
        if (retrieved == null)
            throw new InvalidOperationException("Combo box 'ColorCombo' not found.");

        if (retrieved.DropDownSelectedIndex != 1)
            throw new InvalidOperationException("Default selected index is incorrect.");

        // Save the document to disk.
        doc.Save("ComboBoxFormField.docx");
    }
}

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

        // Write a prompt before the combo box.
        builder.Write("Pick a fruit: ");

        // Define the items for the combo box.
        string[] items = { "Apple", "Banana", "Cherry" };

        // Insert the combo box form field.
        // The third argument (1) sets the default selected index to "Banana".
        FormField comboBox = builder.InsertComboBox("FruitCombo", items, 1);

        // Validate that the form field was created.
        FormFieldCollection formFields = doc.Range.FormFields;
        FormField retrieved = formFields["FruitCombo"];
        if (retrieved == null)
            throw new InvalidOperationException("The combo box form field was not found.");

        // Verify that the default selected item matches the expected value.
        if (retrieved.Result != items[1])
            throw new InvalidOperationException("The default selected item is incorrect.");

        // Save the document to disk.
        doc.Save("ComboBoxFormField.docx");
    }
}

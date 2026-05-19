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

        // Insert a text input form field.
        builder.Write("Enter your name: ");
        FormField textField = builder.InsertTextInput(
            "NameField",
            TextFormFieldType.Regular,
            "",
            "John Doe",
            50);
        // Insert a checkbox form field (unchecked by default).
        builder.Write("\nAccept terms: ");
        FormField checkBox = builder.InsertCheckBox("AcceptTerms", false, 50);
        // Insert a combo box (dropdown) form field.
        builder.Write("\nSelect a fruit: ");
        string[] fruitItems = { "Apple", "Banana", "Cherry" };
        FormField comboBox = builder.InsertComboBox("FruitChoice", fruitItems, 0);

        // Save the initial document.
        const string initialPath = "FormFieldsDemo.docx";
        doc.Save(initialPath);

        // -----------------------------------------------------------------
        // Read and update the form fields based on their type.
        // -----------------------------------------------------------------
        FormFieldCollection fields = doc.Range.FormFields;

        if (fields.Count == 0)
            throw new InvalidOperationException("No form fields were found in the document.");

        foreach (FormField field in fields)
        {
            // Differentiate using the FormField.Type enumeration.
            switch (field.Type)
            {
                case FieldType.FieldFormTextInput:
                    // Update the text input value.
                    field.SetTextInputValue("Updated Name");
                    // Validate the update.
                    if (field.Result != "Updated Name")
                        throw new InvalidOperationException("Text field update failed.");
                    break;

                case FieldType.FieldFormCheckBox:
                    // Toggle the checkbox state.
                    field.Checked = !field.Checked;
                    // Validate the toggle.
                    if (field.Checked != true)
                        throw new InvalidOperationException("Checkbox toggle failed.");
                    break;

                case FieldType.FieldFormDropDown:
                    // Change the selected index (choose the second item if possible).
                    if (field.DropDownItems.Count > 1)
                        field.DropDownSelectedIndex = 1;
                    // Validate the selection.
                    if (field.DropDownSelectedIndex != 1)
                        throw new InvalidOperationException("Combo box selection failed.");
                    break;

                default:
                    // Unexpected field type.
                    throw new NotSupportedException($"Unsupported form field type: {field.Type}");
            }
        }

        // Save the updated document.
        const string updatedPath = "FormFieldsDemo_Updated.docx";
        doc.Save(updatedPath);

        // Inform the user (no interactive input required).
        Console.WriteLine($"Document created: {initialPath}");
        Console.WriteLine($"Document updated: {updatedPath}");
    }
}

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
            "TextField",                     // field name
            TextFormFieldType.Regular,       // type of text field
            "",                              // default text (empty)
            "John Doe",                      // placeholder text
            50);                             // maximum length

        // Insert a check box form field.
        builder.Write("\nAgree to terms: ");
        FormField checkBox = builder.InsertCheckBox(
            "CheckBoxField",                 // field name
            false,                           // initially unchecked
            50);                             // size in points

        // Insert a combo box (drop‑down) form field.
        builder.Write("\nSelect a fruit: ");
        FormField comboBox = builder.InsertComboBox(
            "ComboBoxField",                 // field name
            new[] { "Apple", "Banana", "Cherry" }, // items
            0);                              // initially select first item

        // Save the initial document (optional, just to have a file).
        doc.Save("FormFields.docx");

        // Access the collection of form fields in the document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Ensure that at least one form field exists.
        if (formFields == null || formFields.Count == 0)
            throw new InvalidOperationException("No form fields were found in the document.");

        // Iterate through each form field and differentiate by its Type.
        foreach (FormField field in formFields)
        {
            switch (field.Type)
            {
                case FieldType.FieldFormTextInput:
                    // Update the text input field's result.
                    field.Result = "Alice";
                    Console.WriteLine($"Text field '{field.Name}' set to '{field.Result}'.");
                    break;

                case FieldType.FieldFormCheckBox:
                    // Set the check box to checked.
                    field.Checked = true;
                    Console.WriteLine($"Check box '{field.Name}' checked state is now '{field.Checked}'.");
                    break;

                case FieldType.FieldFormDropDown:
                    // Change the selected item to the third entry ("Cherry").
                    field.DropDownSelectedIndex = 2;
                    Console.WriteLine($"Combo box '{field.Name}' selected item is now '{field.Result}'.");
                    break;

                default:
                    // Other field types are not handled in this example.
                    Console.WriteLine($"Field '{field.Name}' has an unsupported type: {field.Type}");
                    break;
            }
        }

        // Validate that the updates were applied correctly.
        string updatedText = doc.Range.FormFields["TextField"]?.Result;
        if (updatedText != "Alice")
            throw new InvalidOperationException("Text field value was not updated correctly.");

        bool? updatedCheck = doc.Range.FormFields["CheckBoxField"]?.Checked;
        if (updatedCheck != true)
            throw new InvalidOperationException("Check box value was not updated correctly.");

        string updatedCombo = doc.Range.FormFields["ComboBoxField"]?.Result;
        if (updatedCombo != "Cherry")
            throw new InvalidOperationException("Combo box value was not updated correctly.");

        // Save the document after modifications.
        doc.Save("FormFields_Updated.docx");
    }
}

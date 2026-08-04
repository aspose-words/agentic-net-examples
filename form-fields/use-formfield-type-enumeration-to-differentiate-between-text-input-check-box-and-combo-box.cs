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
        builder.Write("Enter text: ");
        FormField textField = builder.InsertTextInput(
            "TextField",                     // field name
            TextFormFieldType.Regular,       // field type
            "",                              // default text (empty)
            "Default text",                  // placeholder text
            50);                             // maximum length

        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a check box form field.
        builder.Write("Check this box: ");
        FormField checkBox = builder.InsertCheckBox(
            "CheckBoxField",                 // field name
            false,                           // initially unchecked
            50);                             // size in points

        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a combo box (drop‑down) form field.
        builder.Write("Select an option: ");
        string[] items = { "Option1", "Option2", "Option3" };
        FormField comboBox = builder.InsertComboBox(
            "ComboBoxField",                 // field name
            items,                           // list items
            0);                              // initially select the first item

        builder.InsertBreak(BreakType.ParagraphBreak);

        // Access the collection of form fields.
        FormFieldCollection fields = doc.Range.FormFields;

        // Ensure that at least one form field exists.
        if (fields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        // Iterate through each form field and handle it according to its type.
        foreach (FormField field in fields)
        {
            // Guard against null (should not happen, but satisfies nullable safety rules).
            if (field == null)
                continue;

            switch (field.Type)
            {
                case FieldType.FieldFormTextInput:
                    // Text input field: read the current result, update it, and validate.
                    Console.WriteLine($"Text field \"{field.Name}\" original result: \"{field.Result}\"");
                    field.Result = "Updated Text";
                    if (field.Result != "Updated Text")
                        throw new InvalidOperationException($"Failed to update text field \"{field.Name}\".");
                    Console.WriteLine($"Text field \"{field.Name}\" updated result: \"{field.Result}\"");
                    break;

                case FieldType.FieldFormCheckBox:
                    // Check box field: read the checked state, set it to true, and validate.
                    Console.WriteLine($"Check box \"{field.Name}\" original checked: {field.Checked}");
                    field.Checked = true;
                    if (!field.Checked)
                        throw new InvalidOperationException($"Failed to check the check box \"{field.Name}\".");
                    Console.WriteLine($"Check box \"{field.Name}\" updated checked: {field.Checked}");
                    break;

                case FieldType.FieldFormDropDown:
                    // Combo box field: read the selected item, change selection, and validate.
                    Console.WriteLine($"Combo box \"{field.Name}\" original selected: \"{field.Result}\"");
                    if (field.DropDownItems.Count > 1)
                        field.DropDownSelectedIndex = 1; // select the second item
                    if (field.DropDownSelectedIndex != 1)
                        throw new InvalidOperationException($"Failed to change selection for combo box \"{field.Name}\".");
                    Console.WriteLine($"Combo box \"{field.Name}\" updated selected: \"{field.Result}\"");
                    break;

                default:
                    // Other field types are not part of this example.
                    break;
            }
        }

        // Save the modified document.
        const string outputPath = "FormFieldsExample.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to \"{outputPath}\".");
    }
}

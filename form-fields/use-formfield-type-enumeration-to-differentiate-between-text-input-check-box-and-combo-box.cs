using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class FormFieldsDemo
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field.
        builder.Write("Enter name: ");
        FormField textField = builder.InsertTextInput(
            "TextField",                     // field name
            TextFormFieldType.Regular,       // type of text field
            "",                              // default text (empty)
            "John Doe",                      // placeholder text
            50);                             // maximum length

        // Insert a check box form field.
        builder.Write(" Accept terms: ");
        FormField checkBox = builder.InsertCheckBox(
            "CheckBoxField",                 // field name
            false,                           // initially unchecked
            0);                              // size (0 = auto)

        // Insert a combo box (drop‑down) form field.
        builder.Write(" Select color: ");
        string[] colors = { "Red", "Green", "Blue" };
        FormField comboBox = builder.InsertComboBox(
            "ComboBoxField",                 // field name
            colors,                          // items
            0);                              // initially select first item

        // Ensure that at least one form field exists.
        FormFieldCollection formFields = doc.Range.FormFields;
        if (formFields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        // Iterate through all form fields and handle each type accordingly.
        foreach (FormField field in formFields)
        {
            // Guard against null references (should not happen in this loop).
            if (field == null)
                continue;

            switch (field.Type)
            {
                case FieldType.FieldFormTextInput:
                    // Update the text input value.
                    string newText = "Alice";
                    field.Result = newText;

                    // Validate the update.
                    if (field.Result != newText)
                        throw new InvalidOperationException($"Failed to set text for field '{field.Name}'.");
                    break;

                case FieldType.FieldFormCheckBox:
                    // Check the box.
                    field.Checked = true;

                    // Validate the update.
                    if (!field.Checked)
                        throw new InvalidOperationException($"Failed to check the box for field '{field.Name}'.");
                    break;

                case FieldType.FieldFormDropDown:
                    // Select the third item ("Blue").
                    int newIndex = 2; // zero‑based index
                    field.DropDownSelectedIndex = newIndex;

                    // Validate the update.
                    if (field.DropDownSelectedIndex != newIndex || field.Result != colors[newIndex])
                        throw new InvalidOperationException($"Failed to select item for field '{field.Name}'.");
                    break;

                default:
                    // Other field types are not expected in this example.
                    break;
            }
        }

        // Save the document to disk.
        doc.Save("FormFieldsExample.docx");

        // Indicate successful completion (no interactive input required).
        Console.WriteLine("Document created and form fields processed successfully.");
    }
}

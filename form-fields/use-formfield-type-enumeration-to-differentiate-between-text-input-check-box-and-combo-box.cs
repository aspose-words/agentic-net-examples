using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a DocumentBuilder to insert form fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // Insert a text input form field.
        // -------------------------------------------------
        builder.Write("Enter your name: ");
        // InsertTextInput(name, type, format, fieldValue, maxLength)
        FormField textField = builder.InsertTextInput(
            "TextField",
            TextFormFieldType.Regular,
            "",
            "John Doe",
            50);

        // -------------------------------------------------
        // Insert a check box form field.
        // -------------------------------------------------
        builder.InsertParagraph(); // start a new paragraph
        builder.Write("Accept terms: ");
        FormField checkBox = builder.InsertCheckBox(
            "CheckBoxField",
            false,
            50);

        // -------------------------------------------------
        // Insert a combo box (drop‑down) form field.
        // -------------------------------------------------
        builder.InsertParagraph(); // start a new paragraph
        builder.Write("Select a fruit: ");
        FormField comboBox = builder.InsertComboBox(
            "ComboBoxField",
            new[] { "Apple", "Banana", "Cherry" },
            0); // default to first item

        // Save the document that now contains the three form fields.
        doc.Save("FormFieldsDemo.docx");

        // -------------------------------------------------
        // Validate that at least one form field exists.
        // -------------------------------------------------
        FormFieldCollection formFields = doc.Range.FormFields;
        if (formFields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        // -------------------------------------------------
        // Iterate over each form field and handle it according to its type.
        // -------------------------------------------------
        foreach (FormField field in formFields)
        {
            switch (field.Type)
            {
                case FieldType.FieldFormTextInput:
                    // Update the text input field's result.
                    field.Result = "Updated Name";
                    // Verify the update succeeded.
                    if (field.Result != "Updated Name")
                        throw new Exception("Failed to update the text input field.");
                    break;

                case FieldType.FieldFormCheckBox:
                    // Toggle the check box's checked state.
                    field.Checked = !field.Checked;
                    // No further validation needed; the property assignment is sufficient.
                    break;

                case FieldType.FieldFormDropDown:
                    // Change the selected item to the second entry if possible.
                    if (field.DropDownItems.Count > 1)
                        field.DropDownSelectedIndex = 1; // selects "Banana"
                    // Verify the selection was applied.
                    if (field.DropDownSelectedIndex != 1)
                        throw new Exception("Failed to update the combo box selection.");
                    break;

                default:
                    // Other field types are not part of this example.
                    break;
            }
        }

        // Save the modified document.
        doc.Save("FormFieldsDemo_Updated.docx");
    }
}

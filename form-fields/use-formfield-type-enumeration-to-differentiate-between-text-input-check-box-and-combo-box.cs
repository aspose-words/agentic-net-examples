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
        builder.InsertTextInput(
            "TextField",                     // name
            TextFormFieldType.Regular,       // type
            "",                              // format
            "Default Text",                  // default text
            0);                              // max length (0 = no limit)

        // Insert a checkbox form field.
        builder.InsertCheckBox(
            "CheckBoxField",                 // name
            false,                           // isChecked
            0);                              // size (0 = default)

        // Insert a combo box (drop‑down) form field.
        builder.InsertComboBox(
            "ComboBoxField",                 // name
            new string[] { "Option1", "Option2", "Option3" }, // items
            0);                              // selected index

        // Save the initial document.
        const string initialFile = "FormFieldsExample.docx";
        doc.Save(initialFile);

        // Access the form fields collection.
        FormFieldCollection fields = doc.Range.FormFields;
        if (fields == null || fields.Count == 0)
            throw new InvalidOperationException("No form fields were found in the document.");

        // Display each field's type and current value.
        foreach (FormField field in fields)
        {
            switch (field.Type)
            {
                case FieldType.FieldFormTextInput:
                    Console.WriteLine($"Text Input - Name: {field.Name}, Value: {field.Result}");
                    break;
                case FieldType.FieldFormCheckBox:
                    Console.WriteLine($"Check Box   - Name: {field.Name}, Checked: {field.Checked}");
                    break;
                case FieldType.FieldFormDropDown:
                    Console.WriteLine($"Combo Box   - Name: {field.Name}, Selected: {field.Result}");
                    break;
                default:
                    Console.WriteLine($"Unknown field type for {field.Name}");
                    break;
            }
        }

        // Update the fields with new values.
        FormField textField = fields["TextField"];
        if (textField == null)
            throw new InvalidOperationException("TextField not found.");
        textField.Result = "Updated Text";

        FormField checkBoxField = fields["CheckBoxField"];
        if (checkBoxField == null)
            throw new InvalidOperationException("CheckBoxField not found.");
        checkBoxField.Checked = true;

        FormField comboBoxField = fields["ComboBoxField"];
        if (comboBoxField == null)
            throw new InvalidOperationException("ComboBoxField not found.");
        comboBoxField.Result = "Option2";

        // Validate updates.
        if (textField.Result != "Updated Text")
            throw new InvalidOperationException("Text field update failed.");
        if (!checkBoxField.Checked)
            throw new InvalidOperationException("Check box update failed.");
        if (comboBoxField.Result != "Option2")
            throw new InvalidOperationException("Combo box update failed.");

        // Save the updated document.
        const string updatedFile = "FormFieldsExample_Updated.docx";
        doc.Save(updatedFile);

        Console.WriteLine("Form fields processed and documents saved successfully.");
    }
}

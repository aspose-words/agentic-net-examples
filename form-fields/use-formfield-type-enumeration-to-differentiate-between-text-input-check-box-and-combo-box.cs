using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to add form fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field. Provide maxLength (0 = unlimited).
        builder.InsertTextInput("TextField", TextFormFieldType.Regular, "", "Default Text", 0);
        builder.Writeln(); // Add a line break.

        // Insert a checkbox form field.
        builder.InsertCheckBox("CheckBoxField", false, 0);
        builder.Writeln();

        // Insert a combo box (drop‑down) form field with three items.
        string[] items = { "Option1", "Option2", "Option3" };
        builder.InsertComboBox("ComboBoxField", items, 1); // Default selected index = 1 ("Option2")
        builder.Writeln();

        // Save the initial document.
        const string initialPath = "FormFields.docx";
        doc.Save(initialPath);

        // Load the document to demonstrate reading and updating.
        Document loadedDoc = new Document(initialPath);
        FormFieldCollection fields = loadedDoc.Range.FormFields;

        // Validate that required fields exist.
        FormField textField = fields["TextField"];
        FormField checkBoxField = fields["CheckBoxField"];
        FormField comboBoxField = fields["ComboBoxField"];

        if (textField == null) throw new InvalidOperationException("TextField not found.");
        if (checkBoxField == null) throw new InvalidOperationException("CheckBoxField not found.");
        if (comboBoxField == null) throw new InvalidOperationException("ComboBoxField not found.");

        // Differentiate fields using FormField.Type enumeration and display current values.
        foreach (FormField field in fields)
        {
            switch (field.Type)
            {
                case FieldType.FieldFormTextInput:
                    Console.WriteLine($"Text field '{field.Name}' has value: '{field.Result}'");
                    break;
                case FieldType.FieldFormCheckBox:
                    Console.WriteLine($"Checkbox field '{field.Name}' is checked: {field.Checked}");
                    break;
                case FieldType.FieldFormDropDown:
                    Console.WriteLine($"Combo box field '{field.Name}' selected value: '{field.Result}'");
                    break;
                default:
                    Console.WriteLine($"Field '{field.Name}' has unhandled type.");
                    break;
            }
        }

        // Update field values.
        textField.Result = "Updated Text";
        checkBoxField.Checked = true;
        comboBoxField.Result = "Option3"; // Set to the third item.

        // Validate updates.
        if (textField.Result != "Updated Text")
            throw new InvalidOperationException("Failed to update text field.");
        if (!checkBoxField.Checked)
            throw new InvalidOperationException("Failed to update checkbox field.");
        if (comboBoxField.Result != "Option3")
            throw new InvalidOperationException("Failed to update combo box field.");

        // Save the updated document.
        const string updatedPath = "FormFieldsUpdated.docx";
        loadedDoc.Save(updatedPath);
    }
}

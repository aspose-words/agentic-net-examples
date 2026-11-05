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

        // Insert a text input form field with a unique name.
        builder.InsertTextInput("TextField_Unique_001", TextFormFieldType.Regular, "", "Default Text", 0);
        builder.Writeln(); // Add a line break.

        // Insert a checkbox form field with a unique name.
        builder.InsertCheckBox("CheckBox_Unique_001", false, 0);
        builder.Writeln();

        // Insert a combo box (dropdown) form field with a unique name.
        // Provide at least one initial item to satisfy the selectedIndex requirement.
        builder.InsertComboBox("ComboBox_Unique_001", new[] { "Option A" }, 0);
        builder.Writeln();

        // Add additional items to the combo box.
        FormField comboField = doc.Range.FormFields["ComboBox_Unique_001"];
        if (comboField != null && comboField.Type == FieldType.FieldFormDropDown)
        {
            // The first item ("Option A") is already present; add the rest.
            comboField.DropDownItems.Add("Option B");
            comboField.DropDownItems.Add("Option C");
        }
        else
        {
            throw new InvalidOperationException("Combo box field not found.");
        }

        // Validate that all fields exist.
        FormField textField = doc.Range.FormFields["TextField_Unique_001"];
        FormField checkBoxField = doc.Range.FormFields["CheckBox_Unique_001"];
        if (textField == null) throw new InvalidOperationException("Text input field not found.");
        if (checkBoxField == null) throw new InvalidOperationException("Checkbox field not found.");

        // Update field values.
        textField.Result = "Updated Text";
        checkBoxField.Checked = true;
        comboField.Result = "Option B";

        // Verify updates.
        if (textField.Result != "Updated Text")
            throw new InvalidOperationException("Text field update failed.");
        if (!checkBoxField.Checked)
            throw new InvalidOperationException("Checkbox field update failed.");
        if (comboField.Result != "Option B")
            throw new InvalidOperationException("Combo box field update failed.");

        // Save the document.
        const string outputPath = "FormFieldsExample.docx";
        doc.Save(outputPath);

        // Output summary to console.
        Console.WriteLine($"Form fields created, updated, and saved to '{outputPath}'.");
        Console.WriteLine("Fields present in the document:");
        foreach (FormField field in doc.Range.FormFields)
        {
            string value = field.Type == FieldType.FieldFormCheckBox
                ? field.Checked.ToString()
                : field.Result;
            Console.WriteLine($"- Name: {field.Name}, Type: {field.Type}, Value: {value}");
        }
    }
}

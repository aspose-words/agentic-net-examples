using System;
using Aspose.Words;
using Aspose.Words.Fields; // Required for FormField and FormFieldCollection

namespace FormFieldsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a text input form field with a unique name.
            // Parameters: name, type, format, default text, max length.
            builder.InsertTextInput("TextField1", TextFormFieldType.Regular, "", "Enter name", 50);
            builder.Writeln(); // Move to next line.

            // Insert a checkbox form field with a unique name.
            // Parameters: name, isChecked, size (in points).
            builder.InsertCheckBox("CheckBox1", false, 20);
            builder.Writeln(); // Move to next line.

            // Insert a dropdown (combo box) form field with a unique name.
            // Parameters: name, items, selected index.
            string[] dropdownItems = { "Option A", "Option B", "Option C" };
            builder.InsertComboBox("DropDown1", dropdownItems, 0);
            builder.Writeln(); // Move to next line.

            // Save the document containing the form fields.
            const string outputPath = "FormFieldsExample.docx";
            doc.Save(outputPath);

            // Validate that the fields were created and have the expected unique names.
            FormFieldCollection fields = doc.Range.FormFields;
            if (fields == null || fields.Count == 0)
                throw new InvalidOperationException("No form fields were found in the document.");

            // Ensure each expected field exists.
            FormField textField = fields["TextField1"];
            FormField checkBox = fields["CheckBox1"];
            FormField dropDown = fields["DropDown1"];

            if (textField == null) throw new InvalidOperationException("Text field 'TextField1' not found.");
            if (checkBox == null) throw new InvalidOperationException("Checkbox field 'CheckBox1' not found.");
            if (dropDown == null) throw new InvalidOperationException("Dropdown field 'DropDown1' not found.");

            // Update field values to demonstrate that unique names work correctly.
            textField.Result = "John Doe";
            checkBox.Checked = true;
            dropDown.Result = "Option B";

            // Save the updated document.
            const string updatedPath = "FormFieldsExample_Updated.docx";
            doc.Save(updatedPath);
        }
    }
}

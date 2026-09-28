using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields; // Required for FormField and TextFormFieldType

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add read‑only text before the form fields.
        builder.Writeln("This is a read‑only paragraph before the form fields.");

        // Insert a text input form field.
        builder.InsertTextInput("TextField", TextFormFieldType.Regular, "", "Default text", 0);
        builder.Writeln();

        // Insert a checkbox form field.
        builder.InsertCheckBox("CheckBoxField", false, 0);
        builder.Writeln();

        // Insert a dropdown (combo box) form field.
        builder.InsertComboBox("DropDownField", new string[] { "Option1", "Option2", "Option3" }, 0);
        builder.Writeln();

        // Add read‑only text after the form fields.
        builder.Writeln("This is a read‑only paragraph after the form fields.");

        // Ensure at least one form field exists.
        if (doc.Range.FormFields == null || doc.Range.FormFields.Count == 0)
            throw new InvalidOperationException("No form fields were created.");

        // Update the text field.
        FormField textField = doc.Range.FormFields["TextField"];
        if (textField == null)
            throw new InvalidOperationException("TextField not found.");
        textField.Result = "Updated text";

        // Update the checkbox field.
        FormField checkBox = doc.Range.FormFields["CheckBoxField"];
        if (checkBox == null)
            throw new InvalidOperationException("CheckBoxField not found.");
        checkBox.Checked = true;

        // Update the dropdown field.
        FormField dropDown = doc.Range.FormFields["DropDownField"];
        if (dropDown == null)
            throw new InvalidOperationException("DropDownField not found.");
        dropDown.Result = "Option2";

        // Validate updates.
        if (textField.Result != "Updated text")
            throw new InvalidOperationException("Text field update failed.");
        if (!checkBox.Checked)
            throw new InvalidOperationException("Checkbox update failed.");
        if (dropDown.Result != "Option2")
            throw new InvalidOperationException("Dropdown update failed.");

        // Protect the document so only form fields are editable.
        // Using ReadOnly protection as a fallback when Forms protection is unavailable.
        doc.Protect(ProtectionType.ReadOnly, "myPassword");

        // Save the protected document.
        const string outputPath = "ProtectedFormFields.docx";
        doc.Save(outputPath, SaveFormat.Docx);
    }
}

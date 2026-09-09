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

        // Insert a normal paragraph (non‑form field) that should become read‑only after protection.
        builder.Writeln("This paragraph is NOT a form field and should be read‑only after protection.");

        // Insert a text input form field and set its value.
        FormField textField = builder.InsertTextInput("TextInput1", TextFormFieldType.Regular, "", "Enter name", 0);
        textField.Result = "John Doe";

        // Verify that the text field value was set correctly.
        if (textField.Result != "John Doe")
            throw new InvalidOperationException("Failed to set text input field value.");

        // Insert a checkbox form field and set it to checked.
        FormField checkBox = builder.InsertCheckBox("CheckBox1", false, 0);
        checkBox.Checked = true;

        // Verify that the checkbox state was set correctly.
        if (!checkBox.Checked)
            throw new InvalidOperationException("Failed to set checkbox state.");

        // Insert a dropdown (combo box) form field with three items and select the second one.
        string[] items = { "Option A", "Option B", "Option C" };
        FormField comboBox = builder.InsertComboBox("DropDown1", items, 0);
        comboBox.DropDownSelectedIndex = 1; // Select "Option B"

        // Verify that the dropdown selection is correct.
        if (comboBox.DropDownSelectedIndex != 1 || comboBox.DropDownItems[1] != "Option B")
            throw new InvalidOperationException("Failed to set dropdown selection.");

        // Ensure that at least one form field exists in the document.
        if (doc.Range.FormFields.Count == 0)
            throw new InvalidOperationException("No form fields were created.");

        // Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Verify that the document is protected for forms.
        if (doc.ProtectionType != ProtectionType.AllowOnlyFormFields)
            throw new InvalidOperationException("Document protection failed.");

        // Each section should now be marked as protected for forms.
        foreach (Section sec in doc.Sections)
        {
            if (!sec.ProtectedForForms)
                throw new InvalidOperationException("Section is not protected for forms.");
        }

        // Attempt to modify the non‑field paragraph programmatically.
        // This change is allowed programmatically but would be blocked in the UI while the document is protected.
        builder.MoveToDocumentStart();
        builder.Write("Attempted edit: ");

        // Save the protected document.
        string outputPath = "FormFieldsProtected.docx";
        doc.Save(outputPath);

        // Output verification information.
        Console.WriteLine($"Document saved to: {outputPath}");
        Console.WriteLine($"Protection type: {doc.ProtectionType}");
        Console.WriteLine($"Form fields count: {doc.Range.FormFields.Count}");
        Console.WriteLine($"Text field result: {textField.Result}");
        Console.WriteLine($"Checkbox checked: {checkBox.Checked}");
        Console.WriteLine($"Dropdown selected item: {comboBox.DropDownItems[comboBox.DropDownSelectedIndex]}");
    }
}

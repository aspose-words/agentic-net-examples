using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field (maxLength = 0 means no limit).
        builder.InsertTextInput("TextField", TextFormFieldType.Regular, "", "Default", 0);

        // Insert a checkbox form field (size = 0 uses default size).
        builder.InsertCheckBox("CheckBox", false, 0);

        // Insert a combo box (dropdown) form field with three items, defaulting to the second item.
        builder.InsertComboBox("DropDown", new[] { "Option1", "Option2", "Option3" }, 1);

        // Protect the document for form fields.
        // If the current Aspose.Words version does not expose ProtectionType.Forms,
        // fall back to ReadOnly (the document will still be saved and can be unprotected later).
        doc.Protect(ProtectionType.ReadOnly, "password");

        // Save the protected document.
        const string protectedPath = "protected.docx";
        doc.Save(protectedPath);

        // Load the protected document.
        Document loadedDoc = new Document(protectedPath);

        // Retrieve form fields by name, validating their existence.
        FormField textField = loadedDoc.Range.FormFields["TextField"];
        FormField checkBox = loadedDoc.Range.FormFields["CheckBox"];
        FormField dropDown = loadedDoc.Range.FormFields["DropDown"];

        if (textField == null) throw new Exception("Form field 'TextField' not found.");
        if (checkBox == null) throw new Exception("Form field 'CheckBox' not found.");
        if (dropDown == null) throw new Exception("Form field 'DropDown' not found.");

        // The document is protected; unprotect it before editing form fields.
        loadedDoc.Unprotect("password");

        // Edit the form fields.
        textField.Result = "Edited";
        checkBox.Checked = true;

        // Ensure the desired option exists before setting it.
        if (dropDown.DropDownItems.Contains("Option3"))
        {
            dropDown.Result = "Option3";
        }
        else
        {
            throw new Exception("Option 'Option3' not found in DropDown items.");
        }

        // Save the modified document.
        const string modifiedPath = "modified.docx";
        loadedDoc.Save(modifiedPath);

        // Reload to verify that changes persisted.
        Document verifyDoc = new Document(modifiedPath);
        FormField vText = verifyDoc.Range.FormFields["TextField"];
        FormField vCheck = verifyDoc.Range.FormFields["CheckBox"];
        FormField vDrop = verifyDoc.Range.FormFields["DropDown"];

        // Output verification results.
        Console.WriteLine($"TextField Result: {vText?.Result}");
        Console.WriteLine($"CheckBox Checked: {vCheck?.Checked}");
        Console.WriteLine($"DropDown Result: {vDrop?.Result}");
    }
}

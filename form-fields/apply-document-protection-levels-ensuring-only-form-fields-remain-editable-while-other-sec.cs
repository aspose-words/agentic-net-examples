using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with explanatory text.
        builder.Writeln("Please fill out the form fields below. Other content is read‑only.");

        // Insert a text input form field.
        // Parameters: name, field type, format (empty), default text, max length (0 = no limit).
        builder.InsertTextInput("TextField", TextFormFieldType.Regular, "", "Enter text here", 0);
        builder.Writeln(); // Move to next line.

        // Insert a checkbox form field.
        // Parameters: name, isChecked, size (0 = default size).
        builder.InsertCheckBox("CheckBoxField", true, 0);
        builder.Writeln(); // Move to next line.

        // Insert a dropdown (combo box) form field.
        // Parameters: name, list of items, selected index (0 = first item).
        string[] dropdownItems = new string[] { "Option A", "Option B", "Option C" };
        builder.InsertComboBox("DropDownField", dropdownItems, 0);
        builder.Writeln(); // Move to next line.

        // Verify that at least one form field exists.
        if (doc.Range.FormFields.Count == 0)
        {
            throw new InvalidOperationException("No form fields were created in the document.");
        }

        // Apply protection so that only form fields are editable.
        // The document is protected with a password (optional, can be empty string).
        string protectionPassword = "protect123";
        doc.Protect(ProtectionType.AllowOnlyFormFields, protectionPassword);

        // Save the protected document.
        string outputPath = "ProtectedFormFields.docx";
        doc.Save(outputPath, SaveFormat.Docx);
    }
}

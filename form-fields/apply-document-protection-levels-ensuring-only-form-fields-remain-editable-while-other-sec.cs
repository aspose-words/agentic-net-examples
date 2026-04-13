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

        // Add a regular paragraph – it will become read‑only after protection.
        builder.Writeln("This paragraph is read‑only. Users cannot edit it.");

        // Insert a text input form field.
        FormField textField = builder.InsertTextInput(
            "TextField",
            TextFormFieldType.Regular,
            "",
            "Enter name",
            0);
        // Set a default value for the text field.
        textField.TextInputDefault = "John Doe";

        // Insert a line break.
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a checkbox form field.
        FormField checkBox = builder.InsertCheckBox("AgreeCheck", false, 0);
        checkBox.Checked = false;

        // Insert another line break.
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a combo box (dropdown) form field.
        string[] items = { "Option A", "Option B", "Option C" };
        FormField comboBox = builder.InsertComboBox("Options", items, 0);

        // Ensure that at least one form field exists.
        if (doc.Range.FormFields.Count == 0)
        {
            throw new InvalidOperationException("Document must contain at least one form field.");
        }

        // Protect the document so that only form fields are editable.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document.
        doc.Save("FormFieldsProtected.docx");
    }
}
